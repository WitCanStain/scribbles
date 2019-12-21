#!/usr/bin/env python

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
        file_path = input("Please give the complete filepath of the .xls or .xlsx file:\n")
        wb = xlrd.open_workbook(file_path)
        break
    except:
        print("ERROR: Please provide a valid filepath with an existing file.")
        continue



def winning_condition(method, valid_votes, seats):  # this function determines whether a simple majority for av or the droop quota for stv is used

    if method == "av":
        print("Iterating with method='alternative vote'" + ", votes=" + str(valid_votes) + " , and seats=" + str(seats) + ". A simple majority is required.")
        return floor(.50 * int(valid_votes))
    elif method == "stv":
        droop = floor((int(valid_votes) / (int(seats) + 1)) + 1)
        print("Iterating with method='single transferable vote'" + ", votes=" + str(valid_votes) + ", and seats=" + str(seats) + ". The Droop quota: " + str(droop) + " will be used.")
        return droop
    else:
        print("Please provide a valid counting method.")
        winning_condition(method, valid_votes, seats)

def check_valid_votes(votelist):  #checks for valid votes and removes any that aren't
    invalid_votes = 0
    row_count = 1
    for x in votelist:
        candidates = []
        for y in x:
            if y in candidates:
                print('Invalid vote found in row ' + str(row_count) + ": " + str(x))
                votelist.remove(x)
                invalid_votes += 1
                row_count += 1
                break
            candidates.append(y)
        row_count += 1

    print("Invalid votes: " + str(invalid_votes))
    print("Valid votes: " + str(len(votelist)))
    return votelist

def count(sheetno):  # this is the function used to do the actual counting, looping over until all winners are selected

    sheet = wb.sheet_by_index(sheetno)  # initialising the variables and objects to be used
    colsno = sheet.ncols
    rowsno = sheet.nrows
    rows = []
    winners = []
    winner_votes = []
    transferred_votes = {}
    init = 0
    winner_status = False


    for rows_index in range (rowsno):  # # create a list of lists, each list corresponding to a row in the spreadsheet
        templist = []
        for cols_index in range (colsno):
            templist.append(sheet.cell_value(rows_index,cols_index))
        rows.append(templist)

    rows = check_valid_votes(rows)  # remove invalid votes
    valid_votes = len(rows)

    # /home/ruby/Downloads/pythontestfile.xls



    candidates = set()
    for x in rows:
        candidates.add(x[0])
    print("The candidates are: ")  # gives a list of the different candidates
    for item in candidates:
        print(item)



    for i in range(len(candidates)):  # meat of the matter

        tally = []
        for row in rows:
            tally.append(row[i])

        column_tally = Counter(tally)
        print(f"Tally at round {i+1}: ")
        winning_quota = winning_condition(method, valid_votes, seats)
        if transferred_votes == {}:
            print(column_tally)
        else:
            column_tally.update(transferred_votes)
            print(column_tally)

        min_name = min(column_tally, key=column_tally.get)
        min_vote = column_tally[min_name]

        for cand in candidates:
            if int(column_tally[cand]) > winning_quota:  # check if there has been a winner or multiple winners
                winning_quota
                print("Candidate " + cand + " has the necessary amount of votes and wins.")
                winner_status = True
                winners.append(cand)
                winner_votes.append(column_tally[cand])
            if len(winners) < int(seats):
                continue
            else:
                break
        if winner_status == True:
            end_sequence(winners, winner_votes)
            break
        else:
            round_no =+ 1
            print("No candidate has the required votes, moving onto round " + str(round_no + 1))
            print("The candidate with the lowest votes is " + min_name + " with " + str(min_vote) + " votes and is eliminated. Their votes are distributed to the other candidates according to subsequent preferences")

            for row in rows:  # transfer the subsequent preference votes to the next round tally
                if row[i] == min_name:
                    transferred_votes[rows[i][init+1]] += 1 + column_tally[rows[i][init+1]]
                else:
                    pass
            for name in transferred_votes:
                transferred_votes[name] += column_tally[name]
        init += 1

def end_sequence(winners, winner_votes):
    print (str(winners) + " have won with respective vote counts " + str(winner_votes))

for sh in range(wb.nsheets):  # the main program loop, going through all the sheets of the spreadsheet file
    print(f"Now calculating the results for sheet {sh+1} of the workbook.")
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
