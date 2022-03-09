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
