"""

"""

from openpyxl import styles
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

import __utils__ as utils
import __utils__.miscellaneous_functions as misc


def update_workbook(workbook):
    # Initialize the name of the current and previous Analysis worksheets.
    worksheet_name_curr = "Analysis"
    worksheet_name_prev = "Analysis (Old)"
    # If the Old Analysis Worksheet exists
    if utils.worksheet_exists(workbook, worksheet_name_prev):
        # Delete that worksheet.
        utils.delete_worksheet(workbook, worksheet_name_prev)
        pass
    # If the Analysis worksheet exists.
    if utils.worksheet_exists(workbook, worksheet_name_curr):
        # Rename the worksheet to the previously deleted worksheet.
        curr_worksheet = utils.get_worksheet(workbook, worksheet_name_curr)
        utils.update_worksheet_name(curr_worksheet, worksheet_name_prev)
        pass
    # Create a new worksheet with the current worksheet name.
    utils.create_new_worksheet(workbook, worksheet_name_curr, 2)
    # Get the new, empty Analysis worksheet.
    worksheet = utils.get_worksheet(workbook, worksheet_name_curr)
    return worksheet
    pass


def get_fills(data_type):
    gold_fill = styles.PatternFill(
        start_color="FFFF00", end_color="FFFF00", fill_type="solid"
    )
    purple_fill = styles.PatternFill(
        start_color="7030A0", end_color="7030A0", fill_type="solid"
    )
    blue_fill = styles.PatternFill(
        start_color="00B0F0", end_color="00B0F0", fill_type="solid"
    )
    green_fill = styles.PatternFill(
        start_color="66FF66", end_color="66FF66", fill_type="solid"
    )
    value_fill = styles.PatternFill(
        start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
    )
    if data_type == "Talent":
        return [purple_fill, blue_fill, green_fill, value_fill, value_fill]
    else:
        return [gold_fill, purple_fill, blue_fill, green_fill, value_fill,
                value_fill]


def style_analysis_worksheet(worksheet):
    # Initialize column width sizes.
    col_widths = [2.77734375, 6.44140625, 4.6640625, 6.0, 5.5546875,
                  9.88671875, 2.77734375, 4.88671875, 6.44140625, 4.6640625,
                  6.0, 5.5546875, 9.88671875, 2.77734375]
    # Set the column sizes of the Analysis worksheet.
    for i, width in enumerate(col_widths):
        col_letter = get_column_letter(i + 1)
        worksheet.column_dimensions[col_letter].width = width
        pass
    # Initialize a black fill color.
    black_fill = styles.PatternFill(
        start_color="000000", end_color="000000", fill_type="solid"
    )
    # For a set number of rows.
    for i in range(1, 100 + 1):
        # Fill the cell in columns A, G, and N with black.
        utils.get_cell(worksheet, (i, 1)).fill = black_fill
        utils.get_cell(worksheet, (i, 7)).fill = black_fill
        utils.get_cell(worksheet, (i, 14)).fill = black_fill
        # If this is the first row.
        if i == 1:
            # For the first 14 cells.
            for j in range(1, 14 + 1):
                # Fill the cell with black.
                utils.get_cell(worksheet, (1, j)).fill = black_fill
                pass
            pass
        pass
    pass


def add_talent_data(worksheet, talent):
    talent_data = []
    talent_runs = sum(talent.values())
    for drop, occurrence in talent.items():
        p, b, g = drop[1:-1].split("/")
        value = misc.compute_drop_value(drop[1:-1].split("/"))
        probability = occurrence / talent_runs
        talent_data.append([int(p), int(b), int(g), value, probability])
        pass
    for i, row in enumerate(talent_data):
        for j, value in enumerate(row):
            r, c = i + 5, j + 2
            utils.update_cell_value(worksheet, (r, c), value)
            if j == len(row) - 1:
                cell = utils.get_cell(worksheet, (r, c))
                cell.number_format = '0.00%'
                pass
            pass
        pass
    pass


def add_weapon_data(worksheet, weapon):
    weapon_data = []
    weapon_runs = sum(weapon.values())
    for drop, occurrence in weapon.items():
        y, p, b, g = drop[1:-1].split("/")
        value = misc.compute_drop_value(drop[1:-1].split("/"))
        probability = occurrence / weapon_runs
        weapon_data.append([int(y), int(p), int(b), int(g), value,
                            probability])
        pass
    for i, row in enumerate(weapon_data):
        for j, value in enumerate(row):
            r, c = i + 5, j + 8
            utils.update_cell_value(worksheet, (r, c), value)
            if j == len(row) - 1:
                cell = utils.get_cell(worksheet, (r, c))
                cell.number_format = '0.00%'
                pass
            pass
        pass
    pass


def style_talent(worksheet, talent):
    # Merge some cells for the Talent Material Main header.
    utils.merge_cells(worksheet, "B2", "F3")
    # Set the main header title for the Talent Material data.
    utils.update_cell_value(worksheet, "B2", "Talent Materials")
    # Get the Talent Header Cell
    header_talent_cell = utils.get_cell(worksheet, "B2")
    # Set the font size to 24
    header_talent_cell.font = Font(size=24)
    # Center this text in this cell horizontally and vertically.
    header_talent_cell.alignment = Alignment(horizontal="center",
                                             vertical="center")
    # Initialize the sub header for the Talent material data.
    sub_header = ["Purple", "Blue", "Green", "Value", "Probability"]
    # Set the 4-th row as the sub header for the Talent Material data.
    for i, header in enumerate(sub_header):
        cell_loc = f"{get_column_letter(i + 2)}4"
        utils.update_cell_value(worksheet, cell_loc, header)
        pass
    # Initialize the Fonts of the sub headers for the Talent Material data.
    talent_fonts = [Font(b=True, color="FFFFFF"),
                    Font(b=True, color="FFFFFF"),
                    Font(b=True, color="000000"),
                    Font(size=11), Font(size=11)]
    # Apply font styling to the sub header for the Talent Material data.
    for i, talent_font in enumerate(talent_fonts):
        cell_loc = f"{get_column_letter(i + 2)}4"
        cell = utils.get_cell(worksheet, cell_loc)
        cell.font = talent_font
        pass
    # Initialize the fills of the sub headers for the Talent Material data.
    talent_fills = get_fills("Talent")
    # Apply fill styling to the sub header for the Talent Material data.
    for i, talent_fill in enumerate(talent_fills):
        cell_loc = f"{get_column_letter(i + 2)}4"
        cell = utils.get_cell(worksheet, cell_loc)
        cell.fill = talent_fill
        pass
    # Initialize a 3-color scale based on percentage.
    color_rule = ColorScaleRule(
        start_type='min', start_color="80F8696B",
        mid_type='percentile', mid_value=50, mid_color="80FFEB84",
        end_type='max', end_color="8063BE7B"
    )
    # Apply a Color scale to the potential non-empty cells in the probability
    # column.
    start_cell, end_cell = "F5", f"F{4 + len(talent)}"
    utils.apply_color_scale(worksheet, start_cell, end_cell, color_rule)
    pass


def style_weapon(worksheet, weapon):
    # Merge some cells for the Weapon Material Main header.
    utils.merge_cells(worksheet, "H2", "M3")
    # Set the main header title for the Weapon Material data.
    utils.update_cell_value(worksheet, "H2", "Weapon Materials")
    # Get the Weapon Header Cell
    header_weapon_cell = utils.get_cell(worksheet, "H2")
    # Set the font size to 24
    header_weapon_cell.font = Font(size=24)
    # Center this text in this cell horizontally and vertically.
    header_weapon_cell.alignment = Alignment(horizontal="center",
                                             vertical="center")
    # Initialize the sub header for the Weapon material data.
    sub_header = ["Gold", "Purple", "Blue", "Green", "Value", "Probability"]
    # Set the 4-th row as the sub header for the Weapon Material data.
    for i, header in enumerate(sub_header):
        cell_loc = f"{get_column_letter(i + 8)}4"
        utils.update_cell_value(worksheet, cell_loc, header)
        pass
    # Initialize the Fonts of the sub headers for the Weapon Material data.
    weapon_fonts = [Font(b=True, color="000000"),
                    Font(b=True, color="FFFFFF"),
                    Font(b=True, color="FFFFFF"),
                    Font(b=True, color="000000"),
                    Font(size=11), Font(size=11)]
    # Apply font styling to the sub header for the Weapon Material data.
    for i, weapon_font in enumerate(weapon_fonts):
        cell_loc = f"{get_column_letter(i + 8)}4"
        cell = utils.get_cell(worksheet, cell_loc)
        cell.font = weapon_font
        pass
    # Initialize the fills of the sub headers for the Weapon Material data.
    weapon_fills = get_fills("Weapon")
    # Apply fill styling to the sub header for the Weapon Material data.
    for i, weapon_fill in enumerate(weapon_fills):
        cell_loc = f"{get_column_letter(i + 8)}4"
        cell = utils.get_cell(worksheet, cell_loc)
        cell.fill = weapon_fill
        pass
    # Initialize a 3-color scale based on percentage.
    color_rule = ColorScaleRule(
        start_type='min', start_color="80F8696B",
        mid_type='percentile', mid_value=50, mid_color="80FFEB84",
        end_type='max', end_color="8063BE7B"
    )
    # Apply a Color scale to the potential non-empty cells in the probability
    # column.
    start_cell, end_cell = "M5", f"M{4 + len(weapon)}"
    utils.apply_color_scale(worksheet, start_cell, end_cell, color_rule)
    pass


def phase3_main(workbook, talent, weapon):
    # Delete, modify, and create a new sheet
    analysis = update_workbook(workbook)
    # Style the Analysis worksheet
    style_analysis_worksheet(analysis)
    # Style the Talent data
    style_talent(analysis, talent)
    # Style the Weapon data
    style_weapon(analysis, weapon)
    # Add the Talent data
    add_talent_data(analysis, talent)
    # Add the Weapon data
    add_weapon_data(analysis, weapon)
    pass
