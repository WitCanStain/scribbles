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
                print(f'Invalid vote found in row {row_count}: {x}')
                votelist.remove(x)
                invalid_votes += 1
                row_count += 1
                break
            candidates.add(y)
        row_count += 1

    print("Invalid votes: " + str(invalid_votes))
    print("Valid votes: " + str(len(votelist)))

    return votelist

def vote_transfer(method, column_tally, eliminated_names, eliminated_name):  # transfers votes, sums the previous votes with the new ones and returns the result as a dictionary
    transferred_votes = {}

    # if '' in candidates:
    #   candidates.remove('')

    def distribution(surplus=1):
        for row in rows:
            if transfer_check(eliminated_names, eliminated_name, row) == True:
                print("UNIT: name row: " + eliminated_name + " " + str(row))
                next_name = row[row.index(eliminated_name) + 1]
                for name in row:
                    if row[row.index(name) + 1] in eliminated_names:  # if a name after the current name has also been eliminated, then move on to the next one
                        continue
                    else:
                        try:
                            transferred_votes[row[row.index(name) + 1]] += surplus
                        except KeyError:
                            if next_name == '':
                                print("No one to transfer votes to!")
                        break
                try:
                    if method == 'av' or method == 'stv_loser':
                        print(f"UNIT: vote to be transferred to {next_name}")
                    elif method == 'stv_winner':
                        print(f"Surplus fraction {str(surplus_fraction)} to be transferred to {next_name}")
                except:
                    print("No more names, last name: " + row[row.index(eliminated_name)])

                try:
                    print(str(next_name) + ' ' + str(transferred_votes[next_name]))
                except KeyError:
                    if next_name == '':
                        print("KeyError: No one to transfer votes to!")

    for cand in candidates:
        transferred_votes[cand] = 0
        print("UNIT: vote_transfer transferred_votes init: " + str(transferred_votes))
    if method == 'av' or method == 'stv_loser':
        # in case of av and stv process 2, distribute all of the eliminated candidate's votes to the subsequent preference.
        distribution()
    elif method == 'stv_winner':
        print("UNIT: eliminated_name votes: " + str(column_tally[eliminated_name]))
        print(str(winning_condition(method,valid_votes, seats)))
        surplus_fraction = (float(column_tally[eliminated_name]) - winning_condition('stv', valid_votes, seats)) / float(column_tally[eliminated_name])
        print(f"The surplus fraction is: {str(surplus_fraction)}")
        # 1st process: distribute surplus vote to subsequent preference fractionally
        distribution(surplus_fraction)

    for cand in candidates:  # adding the transferred votes to the total tally
        column_tally[cand] += transferred_votes[cand]
        print("UNIT: adding votes to column_tally")
        print(cand)
        print(column_tally[cand])

    return column_tally

def transfer_check(eliminated_names, eliminated_name, row):  # checks whether to transfer votes in a given row
    name_in_row = []
    if eliminated_name in row:
        for name in eliminated_names:
            if name in row and row.index(name) < row.index(eliminated_name):  # check if any of the rejected candidates are in the row and have an index less than the most recently eliminated candidate
                name_in_row.append(name)
        if len(name_in_row) == row.index(eliminated_name) and row.index(eliminated_name) != len(row) - 1:  # check if all the names in the row before the min_name have been eliminated
            return True
        else:
            return False
    else:
        return False
    # conditions to return true: 1. the loser's name is in the row. 2. every element in the row before the loser's name is in min_names
    # 3. if the two previous conditions are true, transfer vote to +1 unless +1 is also in losers, in which case to +2 and so on: to the next person not in min_names
    # the function can be used for either av or stv, max or min names so no if differential necessary

def winning_condition(method, valid_votes, seats):
    if method == 'stv' or method == 'STV':
        droop = floor((int(valid_votes) / (int(seats) + 1)) + 1)
        print("droop is " + str(droop))
        return droop
    elif method == 'av' or method == 'AV':
        majority = floor(int(valid_votes) * 0.50)
        print("winning votes: " + str(majority + 1))
        return majority

def count():  # this is the framework for the counting process for each sheet
    tally = []
    filled_seats = 0
    for row in rows:
        tally.append(row[0])
    column_tally = Counter(tally)  # counts the votes

    for i in range(candidate_no):
        print("UNIT: round " + str(i))
        del column_tally['']
        try:
            for winner in winners:
                candidates.remove(winner)
        except:
            pass
        print(f"Tally at round {i + 1}: ")
        print(column_tally)
        winning_votes = winning_condition(method, valid_votes, seats)

        min_name = min(column_tally, key=column_tally.get)
        min_names.append(min_name)
        min_vote = column_tally[min_name]
        max_name = max(column_tally, key=column_tally.get)
        max_names.append(max_name)
        max_vote = column_tally[max_name]

        for cand in candidates:  # check for winners
            print("UNIT: for cand in candidates")
            if column_tally[cand] > winning_votes:
                print(cand + " has achieved the victory condition and is elected.")
                winners.append(cand)
                winner_votes.append(column_tally[cand])
                filled_seats += 1
            else:
                print(cand + " has not won.")
        if filled_seats < int(seats):  # if there are less winners than seats
            print("UNIT: unfilled seats")

            candidates.remove(min_name)
            winning_votes = winning_condition(method, valid_votes, seats)
            del column_tally[min_name]
            if method == 'av':
                print("The candidate with the lowest votes is " + min_name + " with " + str(min_vote) + " votes and is eliminated. Their votes are distributed to the other candidates according to subsequent preferences")
                vote_transfer(method, column_tally, min_names, min_name)
            elif method == 'stv':
                # 1st process: if there have been any winners, distribute their excess votes to the next candidate in fractional proportion to their total votes
                if max_vote > winning_votes:
                    surplus_votes = max_vote - winning_votes
                try:
                    surplus_votes
                except NameError:
                    print("There has been no winners in this round, moving on to process 2: distributing the losing candidate's votes.")
                    vote_transfer('stv_loser', column_tally, min_names, min_name)
                else:

                    print("The candidate with the highest votes is " + max_name + " with " + str(max_vote) + " votes. Their " + str(surplus_votes) + " surplus votes are distributed fractionally to the subsequent preference candidates.")
                    vote_transfer('stv_winner', column_tally, max_names, max_name)
                    vote_transfer('stv_loser', column_tally, min_names, min_name)


        else:
            end_sequence(winners, winner_votes)
            break

def end_sequence(winners, winner_votes):
    print("The winners are " + str(winners) + " with respective votes " + str(winner_votes))

for sh in range(wb.nsheets):  # the main program loop, iterating through all the sheets in the selected workbook
    print(f"\n********** \n\nNow calculating the results for sheet {sh + 1} of the workbook.")

    sheet = wb.sheet_by_index(sh)  # initialising the variables and objects to be used
    colsno = sheet.ncols
    rowsno = sheet.nrows
    rows = []
    winner_name = ''
    winners = []
    winner_votes = []
    winner_status = False
    min_names = []
    max_names = []
    transferred_votes = {}

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
    if '' in candidates:
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

    count()  # let's roll