from __future__ import print_function

import os
import os.path
from typing import Callable, List, Tuple, Union

import gspread
import numpy as np
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from gspread.utils import a1_range_to_grid_range

from cosmix.fixed_volume_mix import FixedVolumeMix
from cosmix.format import GSHEETS_BANDING_COLORS


def auth_google(google_creds_path):
    """Manages google auth pipeline and saves the auth token at `google_creds_path`."""
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


def get_workbook_and_layout_sheets(
    sheet_url: str | None = None,
    sheet_title: str | None = None,
    creds_directory_paths: str | None = None,
    gcp_service_account_path: str | None = None,
    layout_sheet_name="Layout",
):
    """Returns the gspread workbook and Layout sheet situated at `sheet_url`.

    Args:
    sheet_url: URL of the Layout sheet.

    creds_directory_paths: directory where google auth token should be read/saved. If None, falls back on gcp service accounts which is another auth method.
    gcp_service_account_path: path to the gcp service account json file. If None, falls back on creds_directory_paths which is another auth method.

    layout_sheet_name: name of the sheet containing the Layout.
    """
    if creds_directory_paths is not None:
        creds = auth_google(creds_directory_paths)
        gc = gspread.authorize(creds)
    elif gcp_service_account_path is not None:
        gc = gspread.service_account(filename=gcp_service_account_path)
    else:
        raise RuntimeError("No authentication method specified")

    if sheet_url is not None:
        workbook = gc.open_by_url(sheet_url)
    elif sheet_title is not None:
        workbooks = gc.openall(sheet_title)
        if len(workbooks) > 1:
            raise RuntimeError(f"Multiple workbooks with same title {sheet_title}")
        workbook = workbooks[0]
    else:
        raise RuntimeError("No sheet specified")

    sheet = workbook.worksheet(layout_sheet_name)
    return workbook, sheet


def xl_rowcol_to_cell(row_num, col_num, absolute=False):
    """(row,col) to Google sheets cell ID.
    Example: (0,0) -> "A1"

    If `absolute`: (0,0) -> "$A$1"
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

    if not absolute:
        return col_str + str(row_num)
    return "$" + col_str + "$" + str(row_num)


def get_sheet_all_range(sheet):
    """Returns gsheets range corresponding to all the cells in the sheet."""
    max_row, max_col = (
        sheet._properties["gridProperties"]["rowCount"],
        sheet._properties["gridProperties"]["columnCount"],
    )
    return f"{xl_rowcol_to_cell(0,0)}:{xl_rowcol_to_cell(max_row, max_col)}"


def _create_gsheets_table_aux(
    mix: FixedVolumeMix,
    add_total_line=True,
    columns_default_unit=True,
    coordinates_for_formula: Union[None, Tuple] = None,
):
    species_table = mix.species_table(
        columns_default_unit=columns_default_unit, gsheets_value_and_formats=True
    )

    if coordinates_for_formula is not None:
        i, j = coordinates_for_formula
        total_volume_for_formula = mix.computed_volume().m
        if add_total_line:
            total_volume_for_formula = xl_rowcol_to_cell(
                i + len(species_table), j + 3, absolute=True
            )
        for k in range(len(species_table[1:])):
            species = mix.species_list[k]

            if species.inverse_fraction is not None:
                species_table[k + 1][
                    -1
                ] = f"={total_volume_for_formula}/{species.inverse_fraction}"
            elif species.is_completion:
                sum_start = xl_rowcol_to_cell(i + 1, j + 3)
                sum_end = xl_rowcol_to_cell(i + len(species_table) - 2, j + 3)
                species_table[k + 1][
                    -1
                ] = f"={total_volume_for_formula}-SUM({sum_start}:{sum_end})"
            else:
                species_table[k + 1][
                    -1
                ] = f"={xl_rowcol_to_cell(i+k+1,j+2)}*{total_volume_for_formula}/{xl_rowcol_to_cell(i+k+1,j+1)}"

    table_width = len(species_table[0])
    # full_table = [[mix.mix_name] + [""] * (table_width - 1)] + species_table
    # ^ In fact name will be copied later from layout sheet
    full_table = species_table
    if add_total_line:
        full_table += [
            ["Total"]
            + [""] * (table_width - 2)
            + [mix.computed_volume().to(mix.default_volume_unit).to_tuple()[0]]
        ]

    return full_table


def create_gsheets_table(
    mix: FixedVolumeMix,
    add_total_line: bool = True,
    columns_default_unit=True,
    show_units_when_default_units=False,
    extra_columns=[],
    coordinates_for_formula: Union[None, Tuple] = None,
):
    """
    Transforms a mix into a gsheets table (and returns additional gsheets formatting info)

    Args:
    mix (FixedVolumeMix): The mix to transform to gsheets table.

    add_total_line (bool): Adds a final line with total volume to the table.

    columns_default_unit (bool): Uses the mix's default units in each column rather than custom units per cell.

    show_units_when_default_units (bool): if `columns_default_unit` is True, this flag decides whether to show the default unit in all cells or just in the column header.

    coordinates_for_formula: if given, volume cells will be replaced with formula instead of value.

    Returns:
        Returns a table in gsheets format (row x cols 2D array) and a table of same dimension containing google sheets formatting instructions.
    """
    table_aux = _create_gsheets_table_aux(
        mix, add_total_line, columns_default_unit, coordinates_for_formula
    )

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

    for col in extra_columns:
        for i, row in enumerate(table):
            # Check needed because of extra total line
            if i < len(col):
                val, gsheets_format = col[i]
                row.append(val)
                format_table[i].append(gsheets_format)
            else:
                row.append("")
                format_table[i].append(None)

    return table, format_table


def place_table_on_gsheets(
    sheet,
    table,
    format_table=[],
    top_left_origin=(0, 0),
    banding=True,
    banding_ID=0,
):
    """
    Places a gsheets-ready table (as outputted by `create_gsheets_table`) at `top_left_origin` on the `sheet`.

    Args:
    sheet: gsheets object as given by gspread.

    table: gsheets-ready table (as outputted by `create_gsheets_table`).

    format_table: formatting table (as outputted by `create_gsheets_table`).

    top_left_origin: (row, col) 0-indexed of where to place the table on the sheet.

    banding: whether to use Banding (alternating colors) for the table.

    banding_ID: bandings are uniquely identified by an ID in gsheets.

    Returns:
        Returns a table in gsheets format (row x cols 2D array) and a table of same dimension containing google sheets formatting instructions.
    """

    table_height = len(table)
    table_width = len(table[0])
    row0, col0 = top_left_origin
    sheet_range = f"{xl_rowcol_to_cell(row0,col0)}:{xl_rowcol_to_cell(row0+table_height,col0+table_width)}"

    format_instructions = []
    for i, row in enumerate(format_table):
        for j, val in enumerate(row):
            if val is not None:
                format_instructions.append((f"{xl_rowcol_to_cell(row0+i,col0+j)}", val))

    # Batch formatting requests
    requests = []
    for format_instr in format_instructions:
        grid_range = a1_range_to_grid_range(
            format_instr[0], sheet._properties["sheetId"]
        )
        request = {
            "repeatCell": {
                "range": grid_range,
                "cell": {"userEnteredFormat": format_instr[1]},
                "fields": "userEnteredFormat.numberFormat",
            }
        }

        requests.append(request)

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


def reset_sheet(workbook, sheet, rows=1000, cols=1000):
    """
    Deletes and re-creates the `sheet`. This is very hard reset is needed because of Banding...

    Args:
    workbook: workbook object as given by gspread.

    sheet: gsheets object as given by gspread.

    rows: number of rows of in the sheet.

    cols: number of cols in the sheet.

    Raises:
    It probably raises something if the workbook does not contain the sheet...
    """
    # requests = {"requests": [{"updateCells": {"range": {"sheetId": sheet._properties['sheetId']}, "fields": "*"}}]}
    # res = workbook.batch_update(requests)
    # ^ does not remove bandings
    title = sheet._properties["title"]
    workbook.del_worksheet(sheet)
    return workbook.add_worksheet(title, rows=rows, cols=cols)


def _best_top_left(table):
    """Returns indices so to discard empty rows and cols in top left direction."""
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
    """Get the layout as specified in the `layout_sheet` and placed on `layout_range`."""
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
    mix_parser: Callable[[str], FixedVolumeMix],
    insert_formulas_instead_of_values: bool = False,
    extra_columns_function: Callable[[FixedVolumeMix], List] = lambda x: [],
    max_col_size=4,
    merge_repeats=True,
    add_total_line=True,
    columns_default_unit=True,
    show_units_when_default_units=False,
    layout_range="A1:M9",
    target_sheet_name="Targets",
    overwrite_target_sheet=False,
    empty_row_filler=2,
    column_spacing=1,
    row_spacing=1,
    empty_desc="empty",
    print_mixes=False,
    font_size=10,
    place_titles_above_tables=False,
):
    """
    Creates the targets gsheets by reading the layout on the layout sheet, producing each mix using the `mix_parser`
    and then placing each mix's gsheets table on the Targets sheet.

    Args:
    workbook: workbook object as given by gspread.

    layout_sheet: layout gsheets object as given by gspread.

    mix_parser: a function that, given a description of a mix outpus a `FixedVolumeMix`.

    insert_formulas_instead_of_values: puts formulas for volume columns instead of values.

    extra_columns_function: a function that takes a mix and returns extra gsheets column to add.

    max_col_size: the maximum number of columns of any target's table.

    merge_repeats: should mixes corresponding to the same sample be mixed in the Targets' sheet.

    add_total_line (bool): Adds a final line with total volume to the table.

    columns_default_unit (bool): Uses the mix's default units in each column rather than custom units per cell.

    show_units_when_default_units (bool): if `columns_default_unit` is True, this flag decides whether to show the default unit in all cells or just in the column header.

    layout_range: where is the layout placed on the Layout sheet (NotImplemented).

    target_sheet_name: name for the Targets sheet.

    overwrite_target_sheet (bool): prevents Targets' sheet overwriting is set to False.

    empty_row_filler: how many rows should be used in the Targets sheet for each empty row in the layout.

    column_spacing: empty columns to put in the Targets' sheet for each column in the layout.

    row_spacing: empty rows to put in the Targets' sheet for each row in the layout.

    empty_desc: the name of samples that are empty and should be ignored (useful to capture noise sometimes).

    print_mixes: whether this function should print the mixes each time it is putting them in the sheet.

    font_size: font size to use in the Targets' sheet.

    place_titles_above_tables: whether to put the title of each target above the table or at the same level.
    """

    if layout_range != "A1:M9":
        raise NotImplementedError(
            f"Targets placement algorithm implemented only for the case of default position of layout on layout sheet: `{layout_range}`"
        )

    try:
        targets_sheet = workbook.worksheet(target_sheet_name)
        if not overwrite_target_sheet:
            raise ValueError(
                f"The Targets' sheet `{target_sheet_name}` already exists and `overwrite_target_sheet` is set to False"
            )
    except gspread.WorksheetNotFound as e:
        targets_sheet = workbook.add_worksheet(target_sheet_name, rows=100, cols=100)
    targets_sheet = reset_sheet(workbook, targets_sheet)
    targets_sheet.format(
        get_sheet_all_range(targets_sheet),
        {
            "textFormat": {
                "fontSize": font_size,
            },
        },
    )

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
                or (sample_desc in placed and merge_repeats)
            ):
                if j in not_only_repeats_col or not merge_repeats:
                    current_col += max_col_size + column_spacing
                continue

            empty = False
            mix = mix_parser(sample_desc)
            if merge_repeats and len(layout[sample_desc]) > 1:
                mix.resize(mix.total_target_volume * len(layout[sample_desc]))

            table_position = (
                current_row + (1 if place_titles_above_tables else 0),
                current_col,
            )

            table, format_table = create_gsheets_table(
                mix,
                add_total_line=add_total_line,
                columns_default_unit=columns_default_unit,
                show_units_when_default_units=show_units_when_default_units,
                extra_columns=extra_columns_function(mix),
                coordinates_for_formula=(
                    table_position if insert_formulas_instead_of_values else None
                ),
            )
            print(f"Placing target `{sample_desc}`")
            if print_mixes:
                print(mix)
                print()
            # print(current_col,current_row)
            r, u = place_table_on_gsheets(
                targets_sheet,
                table,
                format_table,
                top_left_origin=table_position,
                banding_ID=k,
            )
            requests += r
            updates.append(u)
            k += 1
            max_table_height = max(
                len(table) + (1 if place_titles_above_tables else 0), max_table_height
            )
            placed[sample_desc] = True

            # place title
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
                    }
                }
            )

            if j in not_only_repeats_col or not merge_repeats:
                current_col += max_col_size + column_spacing

        if i in not_only_repeats_row or not merge_repeats or empty:
            current_row += (
                max_table_height + row_spacing + (0 if not empty else empty_row_filler)
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

    if insert_formulas_instead_of_values:
        # Flag is needed to have Gsheet interpret formulaes correctly
        targets_sheet.batch_update(updates, value_input_option="USER_ENTERED")
    else:
        targets_sheet.batch_update(updates)

    body = {"requests": requests}
    workbook.batch_update(body)

    return targets_sheet
