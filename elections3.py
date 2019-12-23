#!/usr/bin/env python
#/home/ruby/Downloads/pythontestfile2.xls
import sys
sys.path.extend(["/home/ruby/.local/lib/python3.8/site-packages"])
import xlrd
from collections import Counter
from math import floor

"""
This is a script that automates the counting process for JCR elections. 
It can utilise either the Alternative Vote (AV) or Single Transferable Vote (STV) methods.
Please provide the script with Excel files only. Ensure that the file contains only
the raw results, i.e. only rows of names, including RON and excluding any headers or additional data. 
If multiple positions are contested at once, separate the raw results into separate sheets in the workbook.
Please note that the script will not consider any candidate with zero votes in first round.
(c) Ruben Drayton, 2019
"""

while True:  # opening the file we'll be working with
    try:
        file_path = input("Please give the complete filepath of the .xls or .xlsx file, q to quit:\n")
        if file_path == 'q':
            quit()
        wb = xlrd.open_workbook(file_path)
        break
    except:
        print("ERROR: Please provide a valid filepath with an existing file.")
        continue

def check_valid_votes(votelist):  # checks for valid votes and removes any that aren't
    invalid_votes = 0
    row_count = 1
    print("Going through the sheet to check for invalid votes...")
    for x in votelist:
        candidates = set()
        if all(i=='' for i in x):
            print(f"Empty vote detected on row {row_count}, purging.")
            votelist.remove(x)
            invalid_votes += 1
            row_count += 1
        for y in x:
            if y in candidates and y != '':
                print(f'Invalid vote found in row ' + {row_count} + ": " + {x})
                votelist.remove(x)
                invalid_votes += 1
                row_count += 1
                break
            candidates.add(y)
        row_count += 1

    print("Invalid votes: " + str(invalid_votes))
    print("Valid votes: " + str(len(votelist)))

    return votelist

def vote_transfer_av(column_tally, min_names, min_name):  # transfers votes, sums the previous votes with the new ones and returns the result as a dictionary
    transferred_votes = {}
    #if '' in candidates:
    #   candidates.remove('')

    for cand in candidates:
        transferred_votes[cand] = 0
        print("UNIT: vote_transfer_av transferred_votes init: " + str(transferred_votes))

    for row in rows:
        if transfer_check(min_names, min_name, row) == True:
            print("UNIT: name row: " + min_name + " " + str(row))
            for name in row:
                if row[row.index(name) + 1] in min_names:  # if a name after the current name has also been eliminated, then move on to the next one
                    continue
                else:
                    try:
                        transferred_votes[row[row.index(name) + 1]] += 1
                    except KeyError:
                        if row[row.index(min_name) + 1] == '':
                            print("No one to transfer votes to!")
                    break
            try:
                print("UNIT: vote to be transferred to " + row[row.index(min_name) + 1])
            except:
                print("No more names, last name: " + row[row.index(min_name)])

            try:
                print(str(row[row.index(min_name) + 1]) + ' ' + str(transferred_votes[row[row.index(min_name) + 1]]))
            except KeyError:
                if row[row.index(min_name) + 1] == '':
                    print("KeyError: No one to transfer votes to!")

    for cand in candidates:  # adding the transferred votes to the total tally
        column_tally[cand] += transferred_votes[cand]
        print("UNIT: adding votes to column_tally")
        print(cand)
        print(column_tally[cand])

    return column_tally

def transfer_check(min_names, min_name, row):  # checks whether to transfer votes in a given row
    name_in_row = []

    # conditions to return true: 1. the loser's name is in the row. 2. every element in the row before the loser's name is in min_names
    # 3. if the two previous conditions are true, transfer vote to +1 unless +1 is also in losers, in which case to +2 and so on: to the next person not in min_names

    if min_name in row:
        for name in min_names:
            if name in row and row.index(name) < row.index(min_name):  # check if any of the rejected candidates are in the row and have an index less than the most recently eliminated candidate
                name_in_row.append(name)
        if len(name_in_row) == row.index(min_name) and row.index(min_name) != len(row) - 1:  # check if all the names in the row before the min_name have been eliminated
            return True
        else:
            return False
    else:
        return False




def winning_condition(method, valid_votes, seats):
    if method == 'stv' or method == 'STV':
        droop = floor((int(valid_votes) / (int(seats) + 1)) + 1)
        print("droop is " + droop)
        return droop
    elif method == 'av' or method == 'AV':
        majority = floor(int(valid_votes) * 0.50)
        print("winning votes: " + str(majority + 1))
        return majority

def count(sheetno):  # this is the framework for the counting process for each sheet
    tally = []
    filled_seats = 0
    for row in rows:
        tally.append(row[0])
    column_tally = Counter(tally)  # counts the votes

    for i in range(candidate_no):
        print("UNIT: round " + str(i))
        del column_tally['']
        print(f"Tally at round {i + 1}: ")
        print(column_tally)
        for cand in candidates:  # check for winners
            print("UNIT: for cand in candidates")
            if column_tally[cand] > winning_condition(method, valid_votes, seats):
                print(cand + " has achieved the victory condition and is elected.")
                winners.append(cand)
                winner_votes.append(column_tally[cand])
                filled_seats += 1
            else:
                print(cand + " has not won.")
        if filled_seats < int(seats):  # if there are less winners than seats
            print("UNIT: unfilled seats")
            if method == 'av':
                min_name = min(column_tally, key=column_tally.get)
                min_names.append(min_name)
                min_vote = column_tally[min_name]
                print("The candidate with the lowest votes is " + min_name + " with " + str(min_vote) + " votes and is eliminated. Their votes are distributed to the other candidates according to subsequent preferences")


                candidates.remove(min_name)
                del column_tally[min_name]
                vote_transfer_av(column_tally, min_names, min_name)
                column_tally

        else:
            end_sequence(winners, winner_votes)
            break


def end_sequence(winners, winner_votes):
    print("The winners are " + str(winners) + " with respective votes " + str(winner_votes))


for sh in range(wb.nsheets):  # the main program loop, iterating through all the sheets in the selected workbook
    print(f"Now calculating the results for sheet {sh + 1} of the workbook.")

    sheet = wb.sheet_by_index(sh)  # initialising the variables and objects to be used
    colsno = sheet.ncols
    rowsno = sheet.nrows
    rows = []
    winners = []
    winner_votes = []
    min_names = []
    transferred_votes = {}
    winner_status = False


    for rows_index in range(rowsno):  # create a list of lists, each list corresponding to a row in the spreadsheet
        templist = []
        for cols_index in range(colsno):
            templist.append(sheet.cell_value(rows_index, cols_index))
        rows.append(templist)

    rows = check_valid_votes(rows)  # remove invalid votes
    valid_votes = len(rows)

    candidates = set()
    for x in rows:
        if all(i == '' for i in x) != True:
            candidates.add(x[0])
    print("The candidates are: ")  # gives a list of the different candidates
    candidates.remove('')
    for item in candidates:
        print(item)

    candidate_no = len(candidates)

    while True:
        method = input("Please advise whether to use AV or STV for this sheet: \n")
        if method != 'av' and method != 'AV' and method != 'stv' and method != 'STV':
            print("Please enter either 'AV' or 'STV' as a parameter.")
            continue
        break
    while True:
        seats = input("Please advise on the number of contested seats (one seat = av is used, multiple seats = stv is used): \n")
        if seats == 0 or seats == None:
            print("Please enter a valid value.")
            continue
        break

    count(sh)  # let's roll