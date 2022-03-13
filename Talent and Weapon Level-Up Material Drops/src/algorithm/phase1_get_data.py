"""
Author: Rayla Kurosaki

File: phase1_get_data.py

Description: This file contains the functionality to parse through and obtain
             data from a Microsoft Excel Workbook.
"""

import __utils__ as utils


def get_talent_data(workbook):
    """
    Creates a dictionary to store the number of type of talent drops.

    :param workbook: The Microsoft Excel Workbook/Spreadsheet to parse
                     through.
    :return: The dictionary to store the number of type of talent drops,
             sorted in descending order.
    """
    # Get the worksheet that contains the list of talent domain drops.
    worksheet = utils.get_worksheet(workbook, "Talent Drop Data")
    # Initialize a dictionary to store the data from the worksheet.
    talent = {}
    # For each row in the spreadsheet.
    for i, row in enumerate(worksheet.iter_rows()):
        # Ignore the first row (the header row).
        if i > 0:
            # Unpack the row.
            p, b, g = row[0].value, row[1].value, row[2].value
            # Create a string top represent the drop.
            string = f"({p}/{b}/{g})"
            # If the drop is in the dictionary.
            if string in talent:
                # Increment the occurrence for that drop by 1.
                talent[string] += 1
                pass
            # If the drop is not in the dictionary.
            else:
                # Add it as a new entry to the dictionary with an occurrence
                # of one.
                talent[string] = 1
                pass
            pass
        pass
    return utils.get_reverse_sorted_dict(talent)


def get_weapon_data(workbook):
    """
    Creates a dictionary to store the number of type of weapon drops.

    :param workbook: The Microsoft Excel Workbook/Spreadsheet to parse
                     through.
    :return: The dictionary to store the number of type of weapon drops,
             sorted in descending order.
    """
    # Get the worksheet that contains the list of weapon domain drops.
    worksheet = utils.get_worksheet(workbook, "Weapon Drop Data")
    # Initialize a dictionary to store the data from the worksheet.
    weapon = {}
    # For each row in the spreadsheet.
    for i, row in enumerate(worksheet.iter_rows()):
        # Ignore the first row (the header row).
        if i > 0:
            # Unpack the row.
            y, p = row[0].value, row[1].value
            b, g = row[2].value, row[3].value
            # Create a string top represent the drop.
            string = f"({y}/{p}/{b}/{g})"
            # If the drop is in the dictionary.
            if string in weapon:
                # Increment the occurrence for that drop by 1.
                weapon[string] += 1
                pass
            # If the drop is not in the dictionary.
            else:
                # Add it as a new entry to the dictionary with an occurrence
                # of one.
                weapon[string] = 1
                pass
            pass
        pass
    return utils.get_reverse_sorted_dict(weapon)


def phase1_main(workbook):
    """
    The main function to call the functions above to get the Talent and
    Weapon data.

    :param workbook: The Microsoft Excel Workbook/Spreadsheet to parse
                     through.
    :return: The Talent and Weapon data.
    """
    return get_talent_data(workbook), get_weapon_data(workbook)
