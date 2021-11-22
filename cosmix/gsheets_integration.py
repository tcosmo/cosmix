from __future__ import print_function

import os
import os.path
import time

import gspread
import numpy as np
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from cosmix.fixed_volume_mix import FixedVolumeMix
from cosmix.format import GSHEETS_BANDING_COLORS


def auth_google(google_creds_path):
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = None
    if os.path.exists(os.path.join(google_creds_path, "google_token.json")):
        creds = Credentials.from_authorized_user_file(
            os.path.join(google_creds_path, "google_token.json"), SCOPES
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError as e:
                flow = InstalledAppFlow.from_client_secrets_file(
                    os.path.join(google_creds_path, "google_credentials.json"), SCOPES
                )
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(google_creds_path, "google_credentials.json"), SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(os.path.join(google_creds_path, "google_token.json"), "w") as token:
            token.write(creds.to_json())
    return creds


def get_workbook_and_layout_sheets(sheet_url, creds_paths, layout_sheet_name="Layout"):
    creds = auth_google(creds_paths)
    gc = gspread.authorize(creds)
    workbook = gc.open_by_url(sheet_url)
    sheet = workbook.worksheet(layout_sheet_name)
    return workbook, sheet


def xl_rowcol_to_cell(row_num, col_num):
    """(row,col) to Google sheets cell ID.
    Example: (0,0) -> "A1"
    """
    # Removed these 2 lines if your row, col is 1 indexed.
    row_num += 1
    col_num += 1

    col_str = ""

    while col_num:
        remainder = col_num % 26

        if remainder == 0:
            remainder = 26

        # Convert the remainder to a character.
        col_letter = chr(ord("A") + remainder - 1)

        # Accumulate the column letters, right to left.
        col_str = col_letter + col_str

        # Get the next order of magnitude.
        col_num = int((col_num - 1) / 26)

    return col_str + str(row_num)


def _create_gsheet_table_aux(
    mix: FixedVolumeMix, add_total_line=True, columns_default_unit=True
):
    species_table = mix.species_table(
        columns_default_unit=columns_default_unit, gsheet_value_and_formats=True
    )
    table_width = len(species_table[0])
    full_table = [[mix.mix_name] + [""] * (table_width - 1)] + species_table
    if add_total_line:
        full_table += [
            ["Total"]
            + [""] * (table_width - 2)
            + [mix.computed_volume().to(mix.default_volume_unit).to_tuple()[0]]
        ]

    return full_table


def create_gsheet_table(
    mix: FixedVolumeMix,
    add_total_line=True,
    columns_default_unit=True,
    show_units_when_default_units=False,
):
    table_aux = _create_gsheet_table_aux(mix, add_total_line, columns_default_unit)

    table, format_table = [], []
    for row in table_aux:
        table.append([])
        format_table.append([])
        for val in row:
            if isinstance(val, tuple) or isinstance(val, list):
                table[-1].append(val[0])
                if columns_default_unit and not show_units_when_default_units:
                    format_table[-1].append(None)
                else:
                    format_table[-1].append(val[1])
            else:
                table[-1].append(val)
                format_table[-1].append(None)

    return table, format_table


def place_table_on_gsheet(
    sheet,
    table,
    format_table=[],
    top_left_origin=(0, 0),
    banding=True,
    banding_ID=0,
):
    table_height = len(table)
    table_width = len(table[0])
    row0, col0 = top_left_origin
    sheet_range = f"{xl_rowcol_to_cell(row0,col0)}:{xl_rowcol_to_cell(row0+table_height,col0+table_width)}"

    format_instructions = []
    for i, row in enumerate(format_table):
        for j, val in enumerate(row):
            if val is not None:
                format_instructions.append((f"{xl_rowcol_to_cell(row0+i,col0+j)}", val))

    for format_instr in format_instructions:
        sheet.format(
            format_instr[0],
            {"numberFormat": {"type": "NUMBER", "pattern": format_instr[1]}},
        )
        time.sleep(0.200)

    # Batch formatting requests
    requests = []
    if banding:
        request = {
            "addBanding": {
                "bandedRange": {
                    "bandedRangeId": banding_ID + 1,
                    "range": {
                        "sheetId": sheet._properties["sheetId"],
                        "startRowIndex": row0,
                        "endRowIndex": row0 + table_height,
                        "startColumnIndex": col0,
                        "endColumnIndex": col0 + table_width,
                    },
                    "rowProperties": {**GSHEETS_BANDING_COLORS},
                },
            },
        }
        requests.append(request)

    return requests, {"range": sheet_range, "values": table}


def reset_sheet(workbook, sheet):
    # requests = {"requests": [{"updateCells": {"range": {"sheetId": sheet._properties['sheetId']}, "fields": "*"}}]}
    # res = workbook.batch_update(requests)
    # ^ does not remove bandings
    workbook.del_worksheet(sheet)
    return workbook.add_worksheet("Targets", rows=1000, cols=1000)


def _best_top_left(table):
    table_np = np.array(table)

    i0 = 0
    for i in range(len(table)):
        if len(list(filter(lambda x: x is not None, table_np[i, :]))) == 0:
            continue
        i0 = i
        break

    j0 = 0
    for j in range(len(table[0])):
        if len(list(filter(lambda x: x is not None, table_np[:, j]))) == 0:
            continue
        j0 = j
        break

    return i0, j0


def extract_layout_from_layout_sheet(layout_sheet, layout_range="A1:M9"):
    layout_values = layout_sheet.get_values(layout_range)

    layout = {}
    for row in layout_values:
        if row[0] in ["A", "B", "C", "D", "E", "F", "G", "H"]:
            for i in range(12):
                cell_name = row[i + 1]
                if cell_name != "":
                    if cell_name not in layout:
                        layout[cell_name] = []
                    layout[cell_name].append(row[0] + str(i + 1))

    return layout


def create_targets(
    workbook,
    layout_sheet,
    mix_parser,
    max_col_size=4,
    merge_repeats=True,
    layout_range="A1:M9",
    target_sheet_name="Targets",
    empty_line_filler=2,
    column_spacing=1,
    row_spacing=1,
    empty_desc="empty",
    print_mixes=False,
):

    if layout_range != "A1:M9":
        raise NotImplementedError(
            f"Targets placement algorithm implemented only for the case of default position of layout on layout sheet: `{layout_range}`"
        )

    try:
        targets_sheet = workbook.worksheet(target_sheet_name)
    except gspread.WorksheetNotFound as e:
        targets_sheet = workbook.add_worksheet("Targets", rows=100, cols=100)
    targets_sheet = reset_sheet(workbook, targets_sheet)

    layout = extract_layout_from_layout_sheet(layout_sheet, layout_range=layout_range)

    inverse_layout = {}
    for sample in layout:
        for cell in layout[sample]:
            inverse_layout[cell] = sample

    tabled_layout = [[None for _ in range(12)] for _ in range(8)]

    not_only_repeats_row = {}
    not_only_repeats_col = {}
    seen = {}

    for i in range(len(tabled_layout)):
        for j in range(len(tabled_layout[0])):
            cell_name = chr(ord("A") + i) + str(j + 1)
            if cell_name in inverse_layout:
                tabled_layout[i][j] = inverse_layout[cell_name]
                if tabled_layout[i][j] not in seen:
                    seen[tabled_layout[i][j]] = True
                    not_only_repeats_row[i] = True
                    not_only_repeats_col[j] = True

    i0, j0 = _best_top_left(tabled_layout)
    k = 0
    current_row = 0
    requests = []
    updates = []
    placed = {}

    for i in range(i0, len(tabled_layout)):
        max_table_height = 0
        empty = True

        current_col = 0

        for j in range(j0, len(tabled_layout[0])):
            sample_desc = tabled_layout[i][j]
            if (
                sample_desc is None
                or sample_desc == empty_desc
                or sample_desc in placed
            ):
                if j in not_only_repeats_col or not merge_repeats:
                    current_col += max_col_size + column_spacing
                continue

            empty = False
            mix = mix_parser(sample_desc)
            if merge_repeats and len(layout[sample_desc]) > 1:
                mix.resize(mix.total_target_volume * len(layout[sample_desc]))
            table, format_table = create_gsheet_table(mix)
            print(f"Placing target `{sample_desc}`")
            if print_mixes:
                print(mix)
                print()
            # print(current_col,current_row)
            r, u = place_table_on_gsheet(
                targets_sheet,
                table,
                format_table,
                top_left_origin=(current_row, current_col),
                banding_ID=k,
            )
            requests += r
            updates.append(u)
            k += 1
            max_table_height = max(len(table), max_table_height)
            placed[sample_desc] = True

            requests.append(
                {
                    "copyPaste": {
                        "source": {
                            "sheetId": layout_sheet._properties["sheetId"],
                            "startRowIndex": i + 1,
                            "endRowIndex": i + 2,
                            "startColumnIndex": j + 1,
                            "endColumnIndex": j
                            + 2,  # TODO this does nto work when layout_range is not default
                        },
                        "destination": {
                            "sheetId": targets_sheet._properties["sheetId"],
                            "startRowIndex": current_row,
                            "endRowIndex": current_row + 1,
                            "startColumnIndex": current_col,
                            "endColumnIndex": current_col + 1,
                        },
                        "pasteType": "PASTE_FORMAT",
                    }
                }
            )

            if j in not_only_repeats_col or not merge_repeats:
                current_col += max_col_size + column_spacing

        if i in not_only_repeats_row or not merge_repeats or empty:
            current_row += (
                max_table_height + row_spacing + (0 if not empty else empty_line_filler)
            )

    # Auto fit
    requests += [
        {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": targets_sheet._properties["sheetId"],
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 100,
                }
            }
        }
    ]

    targets_sheet.batch_update(updates)

    body = {"requests": requests}
    workbook.batch_update(body)

    return targets_sheet
