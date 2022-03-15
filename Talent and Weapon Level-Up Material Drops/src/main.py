"""
Author: Rayla Kurosaki

File: main.py

Description: This program computes the frequency of each type of drop for the
             Talent and Weapon material drops for the game Genshin Impact.
"""

import sys
from os.path import exists as file_exists

import __utils__ as utils
import algorithm as alg


def main():
    # Hardcode the path from source root.
    path = "../data/talent_weapon_drops.xlsx"
    # Check if the file exists
    if not file_exists(path):
        # Exits the program if the file does not exist.
        print("Place your excel file in the data directory.\n"
              "Make sure it is called \"talent_weapon_drops.xlsx\"")
        sys.exit(0)
        pass
    # Get the Excel Spreadsheet.
    workbook = utils.get_workbook(path)
    # Get the talent and weapon drops.
    talent, weapon = alg.phase1_main(workbook)
    # Print the analysis to a text file.
    alg.phase2_main(talent, weapon)
    # Update the analyzed data in the workbook.
    alg.phase3_main(talent, weapon)
    pass


if __name__ == '__main__':
    main()
    pass
