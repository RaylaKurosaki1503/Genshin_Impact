"""
Author: Rayla Kurosaki

File: main.py

Description: This program computes the frequency of each type of drop for the
             Talent and Weapon material drops for the game Genshin Impact.
"""

import sys
from os.path import exists as file_exists

import __utils__ as utils
from projects.Talent_Weapon_Material_Drop_Analysis import algorithm as alg


def main():
    path = "../__data__/talent_weapon_drops.xlsx"
    if not file_exists(path):
        print("Place your excel file in the data directory.\n"
              "Make sure it is called \"talent_weapon_drops.xlsx\"")
        sys.exit(0)
        pass
    workbook = utils.get_workbook(path)
    talent, weapon = alg.phase1_main(workbook)
    alg.phase2_main(talent, weapon)
    alg.phase3_main(talent, weapon)
    pass


if __name__ == '__main__':
    main()
    pass
