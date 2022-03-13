"""
Author: Rayla Kurosaki

File: miscellaneous_functions.py

Description: This file contains all uncategorizable functions.
"""


def get_reverse_sorted_dict(dict):
    """
    Sorts the dictionary in reverse order according to the keys of the
    original dictionary.

    :param dict: Dictionary to sort.
    :return: The sorted dictionary in reverse order.
    """
    # Initialize a temporary dictionary to store the reversed sorted pairs of
    # the original dictionary.
    reverse_sorted_dict = {}
    # For each key in the reversed sorted original dictionary.
    for key in sorted(dict, reverse=True):
        # Add it to the temporary dictionary.
        reverse_sorted_dict[key] = dict[key]
    return reverse_sorted_dict


def format_num_2(num):
    """
    Format the number to 2 decimal places.

    :param num: The number to format.
    :return: The number formatted to 2 decimal places.
    """
    return float("{:.2f}".format(num))


def compute_drop_value(drop):
    """
    Computes the value of the drop. Each green material is worth 1. Each blue
    material is worth 3 green materials. Each purple material is worth 3 blue
    materials. Each gold material is worth 3 purple materials. The value of
    each drop is computed by computing its value in green materials.

    :param drop: The drop to compute the value of.
    :return: The value of the drop.
    """
    # Case where the drop is a talent material drop.
    if len(drop) == 3:
        p, b, g = drop
        # Set the gold drop to 0 since you cannot get gold talent materials.
        y = 0
    # Case where the drop is a weapon material drop.
    else:
        y, p, b, g = drop
    return int(g) + 3 * int(b) + 9 * int(p) + 27 * int(y)
