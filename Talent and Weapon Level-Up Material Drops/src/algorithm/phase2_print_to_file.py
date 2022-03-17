"""
Author: Rayla Kurosaki

File: phase2_print_to_file.py

Description: This file contains functions to analyze and print the data to a
             text file.
"""

import numpy as np

import copy

import __utils__ as utils
import __utils__.miscellaneous_functions as misc


def get_data_to_print(header, data):
    """
    Gets all the necessary data to print.

    :param header: First row of the table.
    :param data: Desired data to print.
    :return: The necessary data to print.
    """
    num_runs = sum(data.values())
    data_to_print = copy.deepcopy(np.array(header))
    for k, v in data.items():
        val = misc.compute_drop_value(k[1:-1].split("/"))
        lst = np.array([
            [k, val, v, f"{utils.format_num_2(100 * v / num_runs)}%"]
        ])
        data_to_print = np.vstack((data_to_print, lst))
    return data_to_print


def print_analysis(datas):
    """
    Prints out the analysis for the Talent material drops and for the Weapon
    material drops.

    :param datas: A list of data to print.
    """
    sxn_headers = ["Talent Domain Drop Analysis",
                   "Weapon Domain Drop Analysis"]
    table_header = [["Drop", "Value", "Occurrence", "Probability"]]
    with open("../data/talent_weapon_drop_analysis.txt", "w") as f:
        for i, (data_i, sub_header) in enumerate(zip(datas, sxn_headers)):
            data = get_data_to_print(table_header, data_i)
            max_len = utils.get_max_len(data)
            spacing = utils.get_spacing(max_len, sub_header)
            f.write(f"{'=' * spacing} {sub_header} {'=' * spacing}\n")
            utils.print_boundary(f, max_len)
            for j, row in enumerate(data):
                if j == 1:
                    utils.print_separator(f, max_len)
                    pass
                utils.print_row(f, row, max_len)
                pass
            utils.print_boundary(f, max_len)
            if i + 1 < len(datas):
                f.write("\n\n")
            pass
        pass
    pass


def phase2_main(talent, weapon):
    """
    This file calls functions to analyze the Talent and Weapon material data
    sets and prints it out to a text file.

    :param talent: The Talent data set to analyze.
    :param weapon: The Weapon data set to analyze.
    """
    print_analysis([talent, weapon])
    pass
