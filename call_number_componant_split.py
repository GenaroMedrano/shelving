import re
from openpyxl import Workbook, load_workbook

# This file checks through each of the issues that may be in a call number. It takes call numbers one at a time.

def remove_extra_callnumbers(inns):
    outs = inns.split(";")
    return outs[0]


def first_character_must_be_letter(inns):
    outs = False
    if inns[0].isalpha():
        outs = True
    return outs


def make_number_if_possible(inns):
    try:
        outs = float(inns)
    except ValueError:
        outs = inns
    return outs


def remove_spaces(inns):
    outs = []
    try:
        inns.remove('')
        inns.remove(' ')
    except ValueError:
        pass
    for i in inns:
        outs.append(i.replace(' ', ''))
    outs = list(filter(None, outs))
    return outs


def break_call_number_into_parts(inns):
    parts = []
    temp_parts = ""
    last = False
    filtered = inns.replace(' .', '.')
    filtered = filtered.replace('. ', '.')
    for i in filtered:
        if i.isdigit() or i == '.':
            current = True
        else:
            current = False
        if last == current:
            temp_parts += i
            last = current
        else:
            if temp_parts != ' ':
                if temp_parts[-1] == '.':
                    parts.append(temp_parts[:-1])
                else:
                    parts.append(temp_parts)
            temp_parts = i
            last = current
    parts.append(temp_parts)
    return parts


def group_calln_parts(inns):
    parts = []
    parts.append([inns[0], inns[1]])
    count = 2
    while count + 1 <= len(inns) - 1:
        if inns[count] and inns[count + 1]:
            parts.append([inns[count], inns[count + 1]])
            count += 2
    if count != len(inns):
        parts.append([inns[-1]])
    outs = parts
    return outs


def leading_zeros(inns):
    outs = inns.zfill(7)
    return outs


def trailing_zeros(inns):
    if inns.isdigit():
        outs = '{:<07d}'.format(int(inns))
    elif inns[0] == '.':
        outs = '.' + leading_zeros(inns[1:])
    else:
        outs = inns
    return outs


def adding_zero_place_holders(inns):
    outs = inns[0][0]
    if '.' in inns[0][1]:
        temp = inns[0][1].split('.')
        outs += leading_zeros(temp[0]) + '.' + trailing_zeros(temp[1])
    else:
        outs += leading_zeros(inns[0][1])
    count = 1
    while count <= len(inns) - 1:
        if len(inns[count]) > 1:
            outs += ' ' + inns[count][0] + trailing_zeros(inns[count][1])
        else:
            outs += ' ' + inns[count][0]
        count += 1
    return outs


def final_output(inns):
    try:
        start = inns.rstrip()
    except:
        return '0000000'
    
    # Removes extra spacing after the data.
    if first_character_must_be_letter(start):
        step1 = remove_extra_callnumbers(start)
        step2 = break_call_number_into_parts(step1)
        step22 = remove_spaces(step2)
        step3 = group_calln_parts(step22)
        step4 = adding_zero_place_holders(step3)
        outs = step4
    else:
        outs = '00000000'
    return outs


def get_call_letters(inns):
    try:
        start = inns.rstrip()  # Removes extra spacing after the data.
    except:
        return '00000000'
    if first_character_must_be_letter(start):
        step1 = remove_extra_callnumbers(start)
        step2 = break_call_number_into_parts(step1)
        step22 = remove_spaces(step2)
        step3 = group_calln_parts(step22)
        step4 = adding_zero_place_holders(step3)
        outs = step4
    else:
        outs = '00000000'
    return outs