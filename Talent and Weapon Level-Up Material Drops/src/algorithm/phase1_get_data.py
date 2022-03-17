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
    worksheet = utils.get_worksheet(workbook, "Talent Drop Data")
    talent = {}
    for i, row in enumerate(worksheet.iter_rows()):
        if i > 0:
            p, b, g = row[0].value, row[1].value, row[2].value
            string = f"({p}/{b}/{g})"
            if string in talent:
                talent[string] += 1
                pass
            else:
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
    worksheet = utils.get_worksheet(workbook, "Weapon Drop Data")
    weapon = {}
    for i, row in enumerate(worksheet.iter_rows()):
        if i > 0:
            y, p = row[0].value, row[1].value
            b, g = row[2].value, row[3].value
            string = f"({y}/{p}/{b}/{g})"
            if string in weapon:
                weapon[string] += 1
                pass
            else:
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
