#!/usr/bin/python

__all__ = [
    "get_next_column",
    "get_numeric_value",
    "get_cell_value",
    "set_cell_value",
    "get_inputs"
]

# Constants

ALPHABET = list("abcdefghijklmnopqrstuvwxyz".upper())

# Functions

def get_next_column(col: str) -> str:
    global ALPHABET

    if ALPHABET.index(col.upper()) + 1 < len(ALPHABET):
        return col.upper()[:-1] + ALPHABET[ALPHABET.index(col.upper()[-1]) + 1].upper()

    return col.upper()[:-1] + ALPHABET[0].upper()

def get_numeric_value(coordinate: str) -> int:
    global ALPHABET

    if coordinate.isalpha():
        digits = list(coordinate.upper())
        digits.reverse()

        value = ALPHABET.index(digits[0]) + 1
        index = len(ALPHABET)

        for digit in digits[1:]:
            value = value + (index * (ALPHABET.index(digit) + 1))
            index = index * index

        return value
    elif coordinate.isnumeric():
        return int(coordinate)
    else:
        raise Exception("Invalid coordinate. Must be numeric (1), or alphabetic (A). Not both (A1)")

def get_cell_value(worksheet, cell: str):
    if worksheet == None:
        raise TypeError("worksheet must not be NoneType")

    return worksheet[cell].value

def set_cell_value(worksheet, cell: str, value):
    if worksheet == None:
        raise TypeError("worksheet must not be NoneType")

    worksheet[cell].value = value

# Inputs

def get_inputs(worksheet, inputs: object, input_pos: str) -> object:
    if worksheet == None:
        raise TypeError("worksheet must not be NoneType")

    if not hasattr(inputs, "__len__"):
        raise TypeError("inputs must be an array")

    input_values = []

    for item in inputs:
        coordinates = ""

        if item.isalpha():
            coordinates = item.upper() + str(input_pos)
        else:
            coordinates = input_pos.upper() + item

        value = get_cell_value(worksheet=worksheet, cell=coordinates)

        if value == None: break

        input_values.append(str(value).strip())

    return input_values


