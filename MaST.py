# -*- coding: utf-8 -*-
"""
Created on Wed Jun 19 23:15:50 2024

@author: D. Brandon Magers

This script aims to replace the entirety of the fortran code formally used to
grade and manange all the students, schools, and individual student test scores
in the MC Math & Science Tournament.

Three Excel/CSV files and one ascii file are needed to run this script

> A key to grade the tests
> Individual student registrations
> School information
> ascii file generated from Wiggin's script from the scantron reader'

run 'python MaST.py --help for general usage information'

Requires 'pip install fpdf2 xlsxwriter'

MY TASKS
> are students told to bubble 05 instead of just 5?
> do we need to figure out rooms before we have the test keys? cause it will try and load a file that doesn't exist yet'

"""

import argparse
import pandas as pd
from datetime import datetime
from time import sleep
import sys
import os
import logging
from colorama import just_fix_windows_console#, init
from fpdf import FPDF
from decimal import Decimal
import ast

# set global variables
sleep_time = 0
num_students_per_school = 12
max_school_id = 425
tests = {'1':'Biology', '2':'Chemistry', '3':'Mathematics', '4':'Physics', '5':'Computer Science'}
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.max_colwidth', 30)

# initialize colors for terminal
just_fix_windows_console()
#init()
def red(phrase): return("\033[91m"+phrase+"\033[0m") 
def grey(phrase): return("\033[90m"+phrase+"\033[0m")
def green(phrase): return("\033[92m"+phrase+"\033[0m")
def purple(phrase): return("\033[95m"+phrase+"\033[0m")
def blue(phrase): return("\033[94m"+phrase+"\033[0m")
def bold(phrase): return("\033[1m"+phrase+"\033[0m")

# welcome message and begin logging file
current_year = datetime.now().year
current_weekday = datetime.now().strftime("%a")
current_month = datetime.now().strftime("%B")
current_day = datetime.now().day
file_date_tag = '-'+str(current_year)+"-"+current_weekday
logfile = 'MaST-'+str(current_year)+'.log'
logging.basicConfig(filename=logfile, level = logging.INFO, format="%(asctime)s %(message)s")
cwd = os.getcwd()

# function to load test keys, schools, and students excel files
def safe_open_excel(flag, file):
    try:
        df = pd.read_excel(file)
    except:
        print(flag+' '+file.ljust(30)+red('Fail'))
        print('\nExiting')
        logging.info('Tried and failed to load '+file+'. Script will exit.')
        sys.exit()
    else:
        print(flag+' '+file.ljust(30)+green('Success'))
        logging.info('Loaded '+file)
    sleep(sleep_time)
    return df

# semi-redudant function but I'm lazy
def safe_open_csv(flag, file):
    try:
        df = pd.read_csv(file)
    except:
        print(flag+' '+file.ljust(30)+red('Fail'))
        print('\nExiting')
        logging.info('Tried and failed to load '+file+'. Script will exit.')
        sys.exit()
    else:
        print(flag+' '+file.ljust(30)+green('Success'))
        logging.info('Loaded '+file)
    sleep(sleep_time)
    return df

# agg df creation
def make_agg_df(df):
    
    # renames agg quantile columns
    def rename(newname):
        def decorator(f):
            f.__name__ = newname
            return f
        return decorator
    
    # calculates variable quantile
    def q_at(y):
        @rename(f'q{y:0.2f}')
        def q(x):
            return x.quantile(y, interpolation='linear')
        return q
    
    # creates agg dataframe with count, max, and quantile values
    q_options = [0.99, 0.98, 0.97, 0.96, 0.95, 0.94, 0.90, 0.88, 0.80, 0.75, 0.50]
    #q_options_small = [0.98, 0.96, 0.94, 0.85, 0.70, 0.50]
    agg_columns = ['count', 'max']
    for item in q_options:
        agg_columns.append(q_at(item))
    agg_series = {'Score': agg_columns}
    data_df_agg = df.groupby('Test').agg(agg_series)
    return data_df_agg

# -- Update --

def update_main(ascii_file, keys_file, students_file, no_lost):
    
    # load raw ascii data file from Scantron
    print('Reading in files...\n')
    sleep(sleep_time)
    ascii_exist = True # sets that there's a new ascii file to load, analyze, and append
    
    # read in new ascii file if it can be found
    if ascii_file != 'None' and ascii_file != 'none':
        try:
            f = open(ascii_file, 'r')
        except:
            print('-a '+ascii_file.ljust(30)+red('Fail').ljust(18)+'Continuing')
            logging.info('Tried and failed to load '+ascii_file)
            ascii_exist = False
        else:
            data = f.read().splitlines()
            f.close()
            print('-a '+ascii_file.ljust(30)+green('Success'))
            logging.info('Loaded '+ascii_file)
    else:
        print('-a '+ascii_file.ljust(30)+blue('Skipped'))
        logging.info('User specified '"--ascii None"' flag so no ASCII file loaded')
        ascii_exist = False
    sleep(sleep_time)
    
    # load keys and students file
    keys_df = safe_open_excel('-k', keys_file)
    students_df = safe_open_excel('-i', students_file)
    
    # cleanup raw data file - remove unneeded columns, split id num into school id and student id, map number to mult choice letter
    def ascii_cleanup(data):
        # checks if there's a blank for test number and puts in a 0 instead which will flag later for user input
        data = [(line[:50]+'0'+line[51:]) if line[50] == ' ' else line for line in data]
        # checks if there's a blank in the school id or test id and if so replaces with 000 and 00 for each respectively
        data = [(line[:40]+'00000'+line[46:]) if ' ' in line[40:45] else line for line in data]
        data = [line.split(None, 6)[4:] for line in data]
        for i in range(len(data)):
            school_id = data[i][0][0:3]
            student_id = data[i][0][3:5]
            test_id = data[i][1]
            test_id = tests.get(test_id,'None') # to replace test number with name 
            answers = list(data[i][2])
            multchoice = {'1':'A', '2':'B', '3':'C', '4':'D', '5':'E'}
            answers = list(map(lambda char: multchoice.get(char,char), answers))
            data[i] = [int(school_id), int(student_id)] + [test_id] + [answers] + [0]
        return data
    
    # check for improper test id
    def check_and_update_test_ids(df):
        flag_improper = False
        for index, row in df.iterrows():
            student_id = int(row['Student ID'])
            school_id = int(row['School ID'])
            test = row['Test']
            if test not in tests.values():
                flag_improper = True
                print(f"\n\nInvalid Test '{test}' for student {student_id} school {school_id} at index {index}.\n")
                print('  Test -> Type exactly Biology, Chemistry, Computer Science, Mathematics, or Physics\n')
                new_test = input(' > Overwrite Test name: ')
                while new_test not in tests.values():
                    print('\nInvalid input. Try again.\n')
                    new_test = input(' > Overwrite Test name: ')
                df.at[index, 'Test'] = new_test
                logging.info("In ASCII data - updated Test field from '"+str(test)+"' to '"+str(new_test)+"' for student "+str(student_id)+" school "+str(school_id)+" at index "+str(index))
        flag_improper_total.append(flag_improper)
        return df
    
    # check for improper school id
    def check_and_update_school_ids(df):
        flag_improper = False
        for index, row in df.iterrows():
            student_id = int(row['Student ID'])
            school_id = int(row['School ID'])
            if school_id < 100 or school_id > max_school_id:
                flag_improper = True
                if school_id == 0: school_id = 'has blank'
                print(f"\n\nInvalid School ID '{school_id}' for student {student_id} at index {index}. Please enter a valid School ID (100-{max_school_id}). Enter 999 to skip permanently.\n")
                new_school_id = int(input(' > Overwrite School ID: '))           
                while new_school_id < 1 or new_school_id > max_school_id and new_school_id != 999:
                    print(f"\nInvalid input. Please enter a School ID between 100 and {max_school_id}.\n")
                    new_school_id = int(input(' > Overwrite School ID: '))           
                df.at[index, 'School ID'] = new_school_id
                logging.info("In ASCII data - updated School ID field from '"+str(school_id)+"' to '"+str(new_school_id)+"' for student "+str(student_id)+" at index "+str(index))
        flag_improper_total.append(flag_improper)
        return df
    
    # check for improper student id
    def check_and_update_student_ids(df):
        flag_improper = False
        for index, row in df.iterrows():
            student_id = int(row['Student ID'])
            school_id = int(row['School ID'])
            if student_id < 1 or student_id > num_students_per_school:
                flag_improper = True
                if student_id == 0: student_id = 'has blank'
                print(f"\n\nInvalid Student ID '{student_id}' for school {school_id} at index {index}. Please enter a valid Student ID (1-12). Enter 99 to skip permanently.\n")
                new_student_id = int(input(' > Overwrite Student ID: '))           
                while new_student_id < 1 or new_student_id > num_students_per_school and new_student_id != 99:
                    print("\nInvalid input. Please enter a Student ID between 1 and 12.\n")
                    new_student_id = int(input(' > Overwrite Student ID: '))           
                df.at[index, 'Student ID'] = new_student_id
                logging.info("In ASCII data - updated Student ID field from '"+str(student_id)+"' to '"+str(new_student_id)+"' for school "+str(school_id)+" at index "+str(index))
        flag_improper_total.append(flag_improper)
        return df    
 
    # Grade - function to count matches based on the test number
    def grade(row, keys_df):
        test_name = row['Test']
    #    test_name = tests.get(test_number,'None')
        answers = row['Answers']
        keys = keys_df[str(test_name)]
        match_count = sum(1 for answer, key in zip(answers, keys) if answer == key)
        return match_count
    
    # cleanup & check for improper data - calls functions above
    flag_improper_total = []
    if ascii_exist:
        print('\nCleaning ASCII file & checking for improper School and Student IDs...', end=' ')
        sleep(sleep_time)
        data = ascii_cleanup(data)
        data_df = pd.DataFrame(data, columns=['School ID','Student ID','Test','Answers','Score'])
        data_df = check_and_update_student_ids(data_df)
        data_df = check_and_update_school_ids(data_df)
        data_df = check_and_update_test_ids(data_df)
        data_df['Score'] = data_df.apply(lambda row: grade(row, keys_df), axis=1)
        if True not in flag_improper_total: print(green('Done'))
    
    # Add new data to old data batch saved in csv file in same folder
    filename = 'MaST-data-'+str(current_year)+'-'+current_weekday+'.csv'
    if ascii_exist:
        print(f'\nAttempting to add new data from ASCII file to previous data in {filename}...', end=' ')
        sleep(sleep_time)
        try:
            og_data_df = pd.read_csv(filename, index_col=0)
        except:
            print(purple('Fail\n'))
            print(f'** Did not find previous file {filename}. Creating new file. **\n')
            logging.info('Did not find previous file '+filename+'. Created file.')
        else:
            print(green('Done\n'))
            og_data_df['Answers'] = og_data_df['Answers'].apply(ast.literal_eval)
            data_df = pd.concat([og_data_df, data_df], ignore_index=True)
            data_df.to_csv(filename)
            logging.info('Loaded '+filename+'. Appended ASCII data to end of file.')
    else:
        print(f'\nAttemping to load previous data from {filename}...', end=' ')
        sleep(sleep_time)
        try:
            og_data_df = pd.read_csv(filename, index_col=0)
        except:
            print(red('Fail\n'))
            print(f'** Did not find previous file {filename}. **\n')
            print('No data to analyze. Exiting script.')
            logging.info('Did not find previous '+filename+'.')
            logging.info('No data to analyze. Script will exit.')
            sys.exit()
        else:
            print(green('Done\n'))
            og_data_df['Answers'] = og_data_df['Answers'].apply(ast.literal_eval)
            data_df = og_data_df
            logging.info('Loaded '+filename+'. Appended ASCII data to end of file.')

    # check that a student didn't take 3 or more tests
    def find_three_plus_same_student(df):
        cont = False
        print('Checking that one student ID did not take 3 or more tests...', end=' ')
        sleep(sleep_time)
        bool_series = df.groupby(['School ID', 'Student ID']).filter(lambda group: group.shape[0] >= 3)[['School ID', 'Student ID', 'Test', 'Answers']]
        if not bool_series.empty:
            print('\n\n'+red('**')+' The following school-student combinations show up 3 or more times '+red('**')+'\n')
            print(bool_series)
            cont = True
        else: print(green('Done\n'))
        return cont
    
    # after find_three_plus_same_student, checks that a student didn't take the same test twice
    def find_same_student_test(df):
        cont = False
        print('Checking that one student did not take the same test twice...', end=' ')
        sleep(sleep_time)
        bool_series = df.groupby(['School ID', 'Student ID', 'Test']).filter(lambda group: group.shape[0] >= 2)[['School ID', 'Student ID', 'Test', 'Answers']]
        if not bool_series.empty:
            print('\n\n'+red('**')+' The following school-student-test combinations show up 2 or more times '+red('**')+'\n')
            print(bool_series)
            cont = True
        else: print(green('Done\n'))
        return cont
    
    # this is messy but it works
    def update_record(df):
        cont = True
        print('To update a record first enter the index number of the record to update. Enter -1 to continue with no further edits.\n')
        index = int(input(' > Index: '))
        if index != -1:
            print('\nWhich field would you like to update? Enter number\n')
            print('   1 - School ID')
            print('   2 - Student ID')
            print('   3 - Test\n')
            field = int(input(' > Field: '))
            while field < 1 or field > 3:
                print('\nInvalid value. Try again.\n')
                field = int(input(' > Field: '))
            print('\nWhat would you like to update it to?\n')
            if field == 1: print('  School ID -> 3 digit integer')
            elif field == 2: print('  Student ID -> integer from 1 to 12')
            elif field == 3: print('  Test -> Type exactly Biology, Chemistry, Computer Science, Mathematics, or Physics\n')
            value = input(' > New Value: ')
            if field == 1 or field == 2: value = int(value) # could check if value is reasonable
            if field == 3:
                while value not in tests.values():
                    print('\nInvalid test name. Try again.\n')
                    value = input(' > New Value: ')
            if field == 1: column = 'School ID'
            elif field == 2: column = 'Student ID'
            elif field == 3: column = 'Test'
            df.at[index, column] = value
            print('\n** Record updated successfully **\n')
            logging.info('Record updated - Index '+str(index)+' '+column+' updated to '+str(value)+'.')
        else: cont = False
        return cont, df
    
    cont = True
    while cont == True:
        cont = find_three_plus_same_student(data_df)
        if not cont: break
        cont, data_df = update_record(data_df)
    
    cont = True
    while cont == True:
        cont = find_same_student_test(data_df)
        if not cont: break
        cont, data_df = update_record(data_df)
   
    # find lost students - have test but not matching registered student
    # does the list of lost students need to be reprinted each time update_record is called? 
    # If so, structure like while loop above...
    def find_lost_student(df):
        found_one = False
        for index, row in df.iterrows():
            student_id = int(row['Student ID'])
            school_id = int(row['School ID'])
            found = students_df.loc[(students_df['school id'] == school_id) & (students_df['student id'] == student_id)]
            if found.empty:
                if not found_one: print('\n')
                print(purple('**')+f' Test exists, but no student found for school id {school_id} student id {student_id} at index {index}.\n')
                found_one = True
        return found_one
    
    print('Searching for lost students...', end=' ')
    found_one = find_lost_student(data_df)
    if found_one:
        cont = True
        while cont == True:
            cont, data_df = update_record(data_df)
        print('')
    else: print(green('Done\n'))
   
    # find lost tests - have a registered student but missing test for said student
    def find_lost_test(df):
        # Reshape students_df to long format
        students_long_df = pd.melt(students_df, id_vars=['school id', 'student id'], value_vars=['test 1', 'test 2'], var_name='test_type', value_name='Test')
        # Rename columns for consistent naming
        students_long_df.rename(columns={'school id': 'School ID', 'student id': 'Student ID'}, inplace=True)
        # Merge with data_df
        merged_df = pd.merge(students_long_df, df, how='left', on=['School ID', 'Student ID', 'Test'], indicator=True)
        # Check for non-matches
        non_matches_df = merged_df[merged_df['_merge'] == 'left_only']
        non_matches_df = non_matches_df[['School ID', 'Student ID', 'Test']]
        # Display the non-matching rows
        merged_df = pd.merge(non_matches_df, students_df, left_on=['School ID', 'Student ID'], right_on=['school id', 'student id'])      
        if merged_df.empty:
            print(green('Done'))
            print('')
        else:
            print('\n\nHere is a list of all missing tests sorted by School ID\n')
            print(merged_df[['School ID', 'Student ID', 'name', 'Test']])
            print('')

    print('Searching for lost tests...', end=' ')
    if not no_lost:
        find_lost_test(data_df)
    else: print(blue('Skipped\n'))
    
    
    # Now I need to grade things
    print('Regrading all tests and saving to .csv file...', end=' ')
    sleep(sleep_time)
    data_df['Score'] = data_df.apply(lambda row: grade(row, keys_df), axis=1)
    data_df.to_csv(filename)
    logging.info('All tests regraded and written to '+filename)
    print(green('Done\n'))
    
    # put quantile values in data_df for every test
    def determine_q(score):
        if data_df_agg.loc[test, ('Score', 'count')] < 100:
            q_list = [0.98, 0.96, 0.94, 0.88, 0.75, 0.50]
        else: q_list = [0.99, 0.98, 0.97, 0.90, 0.80, 0.50]
    #    q_list = [0.99, 0.98, 0.97, 0.90, 0.80, 0.50] # all of these HAVE to be in q_options
        for q in q_list:
            if score >= data_df_agg.loc[test, ('Score', 'q'+'{:.2f}'.format(q))]: # the format makes sure there's a zero at the end of the float number from q_list
                return Decimal('1.00') - Decimal(str(q))
        return Decimal('0.99')
    
    print('Calculating quantiles and total test count...', end=' ')
    sleep(sleep_time)
    # create agg df
    data_df_agg = make_agg_df(data_df)
    for test in data_df['Test'].unique():
        data_df.loc[data_df['Test'] == test, 'Calc Quantile'] = data_df.loc[data_df['Test'] == test, 'Score'].apply(determine_q)
    print(green('Done\n'))
    print('Current test totals\n')
    test_totals = data_df_agg[('Score', 'count')]
    test_sum = test_totals.sum()
    print_tests = pd.concat([test_totals, pd.Series({'Total': test_sum})])
    print_tests = print_tests.rename('Count')
    print(print_tests.to_markdown())
    print('')        
 
    # copy Calc Quantile column to Award Quantile
    data_df['Award Quantile'] = data_df['Calc Quantile']
    
    # sort data and write final file for user edits
    filename = 'MaST-data-final'+file_date_tag+'.xlsx'
    print(f'Sorting by Test & Score then saving all data to {filename}...', end=' ')
    sleep(sleep_time)
    data_df_sorted = data_df.sort_values(by=['Test', 'Score'], ascending=[True, False])
    data_df_sorted.to_excel(filename)
    logging.info(f'Quantiles computed, data sorted, and data written to {filename}')
    print(green('Done\n'))
    
    # final print
    print("If there are further scantron ascii files to process, rerun this script by\n")
    print("    python MaST.py update --ascii newfilename\n")
    print(f"If all scantron files have been processed, open {filename} and edit\nonly the 'Award Quantile' column to desired 1, 2, 3, and 10 percentiles for awards.\n")
    print("Once you have edited the final Excel file, tally results by running\n")
    print("    python MaST.py results\n")
    print('-- Program completed successfully. --')



# --- Results --- 

def results_main(data_file, students_file, schools_file):
    # general printing
    print('-- Results --\n')

    # load data-final, schools, and students
    data_df_final = safe_open_excel('-d', data_file)
    schools_df = safe_open_excel('-s', schools_file)
    students_df = safe_open_excel('-i', students_file)
    
    # create results folder
    results_folder = 'MaST-results'+file_date_tag
    try:
        os.makedirs(results_folder, exist_ok=True)
        print(f"\nFolder '{results_folder}' created successfully or already exists.\n")
    except Exception as e:
        print(f"An error occurred: {e}")
    
    # points mapping for school points, all of these options HAVE to show up in q_options
    points_mapping_large = {
        '0.01': 10,
        '0.02': 8,
        '0.03': 6,
        '0.10': 4,
        '0.20': 2,
        '0.50': 1,
    }
    points_mapping_small = {
        '0.02': 10,
        '0.04': 8,
        '0.06': 6,
        '0.12': 4,                              
        '0.25': 2,
        '0.50': 1,
    }
    
    # make agg df
    data_df_agg = make_agg_df(data_df_final)
    
    # Function to map points based on subject and quantile
    def assign_school_points(row):
        if data_df_agg.loc[row['Test'], ('Score', 'count')] >= 100:
    ##        return points_mapping_large.get('{:.2f}'.format(row['Calc Quantile']), 0)
            return points_mapping_large.get(str(row['Calc Quantile']), 0)
        elif data_df_agg.loc[row['Test'], ('Score', 'count')] < 100:
    ##        return points_mapping_small.get('{:.2f}'.format(row['Calc Quantile']), 0)
            return points_mapping_small.get(str(row['Calc Quantile']), 0)
        else:
            return 0
    
    # Add a new column "School Points" using the apply method
    print('Assigning School Points, ranking, and saving file...', end=' ')
    sleep(sleep_time)
    data_df_final['School Points'] = data_df_final.apply(assign_school_points, axis=1)
        
    # Sum the total points for each school
    school_points = data_df_final.groupby('School ID')['School Points'].sum()
    sorted_school_points = school_points.sort_values(ascending=False)
    merge_df = pd.merge(sorted_school_points, schools_df, left_on=['School ID'], right_on=['id'])
    merge_df = merge_df[['id', 'School Points', 'school']]
    file_path = os.path.join(results_folder, "MaST-School_Rankings"+file_date_tag+".xlsx")
    merge_df.to_excel(file_path)
    print(green('Done\n'))

    # print top schools
    print('Top 5 Schools\n')
    print('    ID   Points   Name')
    rank = 1
    for index, value in sorted_school_points.head(5).items():
        school_name = schools_df.loc[schools_df['id'] == index, 'school'].values[0]
        print(f' {rank}. {index}  {str(value).rjust(3)}      {school_name}')
        rank += 1
    print('')
    
    
    # now make excel/csv file for printing certificates
    print('Creating .csv files for printing student certificates...', end=' ')
    sleep(sleep_time)
    
    # Define the list of percentiles to create tabs for
    percentiles = [0.01, 0.02, 0.03, 0.10]
    #percentiles = [0.02, 0.04, 0.06, 0.12]
    
    # Group the data by 'Subject'
    tests_unique = data_df_final['Test'].unique()
    
    # Create an Excel file for each subject
    for test in tests_unique:
        # Filter data for the current subject
        test_df = data_df_final[data_df_final['Test'] == test]

        # Create an Excel writer object
        file_path = os.path.join(results_folder, f'MaST-{test}_winners'+file_date_tag+'.xlsx')
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            # Create a sheet for each percentile
            for percentile in percentiles:
                # Filter data for the current percentile
                percentile_df = test_df[test_df['Award Quantile'] == percentile]

                # Create a DataFrame with just the 'School ID' and 'Student ID' columns
                ids_df = percentile_df[['School ID', 'Student ID']]
                merged_df = pd.merge(ids_df, students_df, left_on=['School ID', 'Student ID'], right_on=['school id', 'student id'])
                names_df = merged_df[['name']]
##                names_df.reset_index(drop=True, inplace=True)
                
                # Write the DataFrame to the Excel sheet
                names_df.to_excel(writer, sheet_name=str(percentile*100)+'%', index=False, header=['Name'])
    print(green('Done\n'))
        
    # make printouts for teachers
    print('Creating .pdf file for school result summaries...', end=' ')
    sleep(sleep_time)
    data_df_final_grouped_schoolid = data_df_final.groupby('School ID')
    pdf = FPDF(unit="mm", format="Letter")
    
    # Function to add school data to the PDF
    def add_school_to_pdf(school_id, group, pdf):
        pdf.add_page()
        
        # Set title
        pdf.set_font("Helvetica", size=12)
        pdf.cell(200, 8, text="Mississippi College", new_x="LMARGIN", new_y="NEXT", align='C')
        pdf.cell(200, 8, text="Math & Science Tournament", new_x="LMARGIN", new_y="NEXT", align='C')
        current_date = str(current_month)+' '+str(current_day)+', '+str(current_year)
        pdf.cell(200, 8, text=current_date, new_x="LMARGIN", new_y="NEXT", align='C')
        pdf.ln(4)
        school_name = schools_df.loc[schools_df['id'] == school_id, 'school'].values[0]
        pdf.cell(200, 8, text=f"{school_name} (ID: {school_id})", new_x="LMARGIN", new_y="NEXT", align='C')
        
        # Set column headers
        pdf.set_font("Helvetica", size=10)
        pdf.cell(10, 8, text="ID", border=1, align='C')
        pdf.cell(60, 8, text="Student Name", border=1)
        pdf.cell(60, 8, text="Test", border=1)
        pdf.cell(20, 8, text="Score", border=1, align='C')
        pdf.cell(20, 8, text="Percentile", border=1, align='C')
        pdf.ln()
        
        # Add rows
        for index, row in group.iterrows():
            pdf.cell(10, 8, text=str(row['Student ID']), border=1, align='C')
            pdf.cell(60, 8, text=students_df.loc[(students_df['school id'] == school_id) & (students_df['student id'] == row['Student ID']), 'name'].values[0], border=1)
            pdf.cell(60, 8, text=str(row['Test']), border=1)
            pdf.cell(20, 8, text=str(row['Score']), border=1, align='C')
            pdf.cell(20, 8, text="{:.0f}".format(row['Calc Quantile']*100), border=1, align='C')
            pdf.ln()
    
    # Add each school's data to the PDF
    for school_id, group in data_df_final_grouped_schoolid:
        add_school_to_pdf(school_id, group, pdf)
    
    # Save the PDF
    file_path = os.path.join(results_folder, "MaST-School_Report"+file_date_tag+".pdf")
    pdf.output(file_path)
    print(green('Done\n'))
    
    # final printing
    print(f'** All result files can be found in the {results_folder} folder **\n')
    print('-- Program completed successfully. --')
    

def main():
    # command line argparser
    parser = argparse.ArgumentParser(description= "MC Math & Science Tournament Grader & Analysis.")
    subparsers = parser.add_subparsers(dest='command', help="Specify either 'update' or 'results'.", required=True)

    # sub-parser for update function
    parser_update = subparsers.add_parser('update', help='Update data and generate CSV file.')
    parser_update.add_argument('-a', '--ascii', type=str, default='MaST-Raw.dat', help='Location of ascii file from Scantron.\n(Default: MaST-ascii.dat')
    parser_update.add_argument('-k', '--keys', type=str, default='MaST-Keys.xlsx', help='Location of excel file with test keys.\n(Default: MaST-Keys.xlsx')
    parser_update.add_argument('-i', '--students', default='MaST-Students.xlsx', help='Location of excel file with student registration information.\n(Default: MaST-Students.xlsx')
    parser_update.add_argument('--no_lost', action='store_true', help='Specify to not print missing tests. Helpful at beginning when not all scantron files have been processed.')

    # sub-parser for results parser
    parser_results = subparsers.add_parser('results', help='Read in final csv file and tally results. No new data will be configured.')
    data_filename = 'MaST-data-final'+file_date_tag+'.xlsx'
    parser_results.add_argument('-d', '--data', default=data_filename, help='Location of final datat excel file made with update mode.\n(Default: '+data_filename+')')
    parser_results.add_argument('-i', '--students', default='MaST-Students.xlsx', help='Location of excel file with student registration information.\n(Default: MaST-Students.xlsx')
    parser_results.add_argument('-s', '--schools', default='MaST-Schools.xlsx', help='Location of excel file with school information.\n(Default: MaST-Schools.xlsx')

    args = parser.parse_args()

    print('\n-- MaST.py @author: D. Brandon Magers --\n')
    logging.info('--------------------')
    logging.info('MaST.py '+args.command+' script began in '+cwd+'at '+str(datetime.now()))

    if args.command == 'update':
        update_main(args.ascii, args.keys, args.students, args.no_lost)
    elif args.command == 'results':
        results_main(args.data, args.students, args.schools)
    else:
        parser.print_help()
    
if __name__ == "__main__":
    main()
