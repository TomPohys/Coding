#!/usr/bin/env python
# encoding: utf-8
"""
Created by Tomas Pohanka on 2016-1-5.
Copyright (c) Tomas Pohanka 2015"""

# Input file from ET
TIMING = 'Hypothesis-s4.xls'
# Input file with times and names of pictures
KEYS = 'keys.xls'
# Text in EventParam -> original Key: Down
TEXTOFKEY = "Key: Down"
# variable for difference between time in TIMING and KEYS [ms]
PLUSMINUS = 16
# EventParam -> for example Key: Down
EVENT = 18
# Original time
TIME = 19
# GazePosX - bad character
BADCHAR = 22
# Header
HEADER = 7

import xlrd
import xlwt

print "- Loading KEYS"
keys = {}
book_keys = xlrd.open_workbook(KEYS)
sh = book_keys.sheet_by_index(0)
# rw - row - counter of lines
# sh.nrows - number of rows
# format keys - {time:(value,order)} - {100:(pic1,1)}
for rw in range(sh.nrows):
    if not rw == 0:
        k = sh.row_values(rw)
        keys[k[0]] = k[1], k[2]

print "- Loading TIMING"
start_time = 0
timing = {}
book_timing = xlrd.open_workbook(TIMING)
write_book_timing = xlwt.Workbook(encoding='utf-8')
sh = book_timing.sheet_by_index(0)
write_sh = write_book_timing.add_sheet('Sheet 1', cell_overwrite_ok = True)
first_down = False
name = ""
order = 0
# 1 because in row 0 will be header
new_row = 1

# For every single row
for row in range(sh.nrows):
    # Commentary in TIMING 
    if not sh.row_values(row)[0].startswith("#"):
        # Header
        if not row == HEADER:
            # first key down (start)
            if sh.row_values(row)[EVENT] == TEXTOFKEY and first_down is False:
                start_time = sh.row_values(row)[TIME]
                new_time = int(sh.row_values(row)[TIME]) - int(start_time)
                if new_time in keys:
                    name = keys[new_time][0]
                    order = keys[new_time][1]
                first_down = True
                # Write the original data
                for col in range(sh.ncols):
                    # Coortinate containt unexpected value in second and third place and the type is string or unicode
                    if col == BADCHAR and type(sh.cell(row, col).value) != float:
                        value = sh.cell(row, col).value
                        value = value.replace(value[1:3], "")
                        write_sh.write(new_row, col, value)
                        continue
                    write_sh.write(new_row, col, sh.cell(row, col).value)

                write_sh.write(new_row, sh.ncols, name)
                write_sh.write(new_row, sh.ncols + 1, new_time)
                write_sh.write(new_row, sh.ncols + 2, order)

                new_row = new_row + 1
                continue

            # second key down (end)
            if sh.row_values(row)[EVENT] == TEXTOFKEY and first_down is True:
                stop_time = sh.row_values(row)[TIME]
                end_time = int(sh.row_values(row)[TIME]) - int(start_time)
                first_down = False
                
                for col in range(sh.ncols):
                    if col == BADCHAR and type(sh.cell(row, col).value) != float:
                        value = sh.cell(row, col).value
                        value = value.replace(value[1:3], "")
                        write_sh.write(new_row, col, value)
                        continue
                    write_sh.write(new_row, col, sh.cell(row, col).value)

                write_sh.write(new_row, sh.ncols, name)
                write_sh.write(new_row, sh.ncols + 1, end_time)
                write_sh.write(new_row, sh.ncols + 2, order)

                name = ""
                new_row = new_row + 1
                continue

            # between key down
            if first_down is True and not sh.row_values(row)[EVENT] == TEXTOFKEY:

                middle_time = int(sh.row_values(row)[TIME]) - int(start_time)
                
                for col in range(sh.ncols):
                    if col == BADCHAR and type(sh.cell(row, col).value) != float:
                        value = sh.cell(row, col).value
                        value = value.replace(value[1:3], "")
                        write_sh.write(new_row, col, value)
                        continue
                    write_sh.write(new_row, col, sh.cell(row, col).value)
                    
                if middle_time in keys:
                    name = keys[middle_time][0]
                    order = keys[middle_time][1]
                    write_sh.write(new_row, sh.ncols, name)
                    write_sh.write(new_row, sh.ncols + 2, order)

                else:
                    # Try to find a correct time if time in TIMING and KEYS is not equal
                    for i in xrange(PLUSMINUS):
                        # Bigger original value in TIMING, smaller in KEYS
                        extra_middle_time = middle_time - i
                        if extra_middle_time in keys and keys[extra_middle_time] > middle_time:
                            keys[middle_time] = keys[extra_middle_time]
                            del keys[extra_middle_time]
                            name = keys[middle_time][0]
                            order = keys[middle_time][1]
                        
                        # Bigger time value in KEYS, smaller in TIMING => set next higher value of time
                        elif (middle_time + i) in keys and (middle_time + i) in keys < middle_time:
                            keys[
                                (int(sh.row_values(row + 1)[TIME]) - int(start_time))] = keys[middle_time + i]
                            del keys[middle_time + i]
                            continue
                        else:
                            i = i + 1

                        write_sh.write(new_row, sh.ncols, name)
                        write_sh.write(new_row, sh.ncols + 2, order)

                write_sh.write(new_row, sh.ncols + 1, middle_time)
                new_row = new_row + 1
        else:
            # Write header
            for col in range(sh.ncols):
                write_sh.write(0, col, sh.cell(row, col).value)
            write_sh.write(0, sh.ncols, "Name")
            write_sh.write(0, sh.ncols + 1, "Time2")
            write_sh.write(0, sh.ncols + 2, "Order")

write_book_timing.save("new_" + TIMING)
