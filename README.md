# MaST.py

@author: D. Brandon Magers

This script aims to replace the entirety of the fortran code formally used to grade and manange all the students, schools, and individual student test scores in the MC Math & Science Tournament.

### Required files and libraries

Two additional Python libraries are required: `fpdf2` and `xlsxwriter`. These can easily be installed with pip.

#### User definied files

The following files are required.

- `MaST-Schools.xlsx` Spreadsheet data of all MS schools. Do not edit ID numbers. See sample.
- `MaST-Students.xlsx` Spreadsheet data of all registered students for the current year. See sample.
- `MaST-Keys.xlsx` Spreadsheet of answer key for all tests. See sample.
- `MaST-Raw.dat` ASCII data file produced from Scantron 

## Usage

There is two main usage modes called update and results. The update mode is for adding new scantron files and checking the data for errors. The results mode is for tallying winners and creating files for printing.  Generally, these two modes are completely separate. The results mode should not be run until all scantron files have been processed with the update mode.

### Update mode

Update mode is activated by adding the 'update' keyword to the command line script call.

`python MaST.py update`

In a typical run in this mode, three of the required files listed above are needed: students, keys, and raw.data. Optionally, specific filenames can be specified with command line flags. 

`python MaST.py update --ascii asciifilename --students studentsfilename --keys keysfilesname

#### Update mode workflow

The update mode does the following tasks in this order.

1. Reads in all files
2. Extracts data from raw scantron file and stores in Pandas DataFrame
3. Checks for students ids not in the range from 1-12 and prompts user to update entry
4. Checks for schools ids not in the range from 100-425 and prompts user to update entry
5. Checks for test ids not in the range from 1-5 and prompts user to update entry
6. Grades tests from imported Scantron ASCII file
7. If it exists, reads in data from previous runs stored in MaST-data-YEAR-DAY.csv file. Adds newly processed scantron data to the previous data and rewrites the files. All further steps are on the entire data set of new and formal tests processed thus far.
8. Checks that one student did not take 3 or more tests and prompts user to update entry
9. Checks that one student did not take the same test twice and prompts user to update entry
10. Searches for lost students and prompts user to update entry
11. Searches for lost tests and prints to screen. Output can be supressed with `--no_lost` flag.
12. Overwrites data to MaST-data-YEAR-DAY.csv
13. Calculates totals and quantile values for each test. Prints current totals to screen.
14. Assigns every test to a quantile bucket
15. Writes all current test data to MaST-data-final-YEAR-DAY.xlsx file

#### Update mode usage notes

It is highly recommended that each Scantron ASCII file be given a different name and numbered by batch. For example, MaST-ascii-1.dat, MaST-ascii-2.dat, etc.

There are a few things to note from the update mode workflow. The MaST-data-YEAR-DAY.csv file is essentially the database that stores the cummulative test data from each time the script is run to add an additional Scantron ASCII file. It can be edited, but should be done so cautiously. If edits are made, it would be wise to run the script in update mode with the `--ascii none` flag option once. This flag skips steps 2-6 above and doesn't add any new data tot he existing file, but does run all steps 8-15. This is important because if one test is changed, tests should be regraded and quantiles recalculated.

For steps 3-4 which checks for improper school or student ids, if a blank is found in either entry, the scripts sets both the school and student id number to '000' and '00' respectively, to force the user to check for the correct value.

For step 11 which searches for lost tests, early on when few tests have been processed, this list could be hundreds of lines long. It is recommended to apply the `--no_lost` flag, which only supresses the output of this list.  Once reaching the end(ish) the end of the scantron stack, do not include this flag to see a list of lost tests printed to the screen.

Quantiles are computed in different ranges depending on the total tests for that subject area. For 100+ tests, 1%, 2%, 3%, 10%, 20%, and 50% are computed and categorized. For 99 or less tests, the quantiles are 2%, 4%, 6%, 12%, 25%, and 50%. This is to ensure there is a 1, 2, and 3% winner for each test. See below on editing winner category ranges.

Even though the MaST-data-final-YEAR-DAY.xlsx file is created every run, it generally should not be needed until all Scantron ASCII files are read in. Opening it will reveal the current highest scoring test for each subject area.

When all Scantron ASCII files have been processed, it is recommended to run the script in update mode once with the `--ascii none` flag to finalize any data errors the script can find. Remember if the .csv data file is edited to rerun again with this flag.

#### Transition between update mode to results mode

Once all Scantron ASCII files have been processed. The MaST-data-final-YEAR-DAY.xlsx needs to be opened and inspected before running the results mode.  In this file, the only column that needs to be edited is the 'Award Quantile' column. If you see other corrections or errors, please refer back to previous steps described above. At first open, the 'Award Quantile' column will mirror the 'Calc Quantile' column. The 'Calc Quantile' column is the calculated quantile bucket based on the ranges defined for < 100 or > 100 tests. All tests are sorted by subject and then by score so rank is easy to see. In the 'Award Quantile' column, edit as needed to define 1, 2, 3, and 10 percentile winners. Any test/student that has these four values will be included in the final results tally. For a subject that has less than 100 tests, please take special note. Even though the calculated quantiles are at 2, 4, 6, and 12 (to help you see distinctions), winners are awarded at 1, 2, 3, and 10 no matter what. So ensure that these categories are represented in this final column. Save the final when done.

### Results mode

Results mode is activated by adding the 'results' keyword to the command line script call.

`python MaST.py results`

In a typical run in this mode, two of the required files listed above are needed: schools and students. Additionally, the MaST-data-final-YEAR-DAY.xlsx file created by the update mode is required. Optionally, specific filenames can be specified with command line flags. 

`python MaST.py results --data MaST-data-final-YEAR-DAY.xlsx --students studentsfilename --schools schoolsfilename

#### Results mode workflow

The results mode does the following tasks in this order.

1. Reads in all files
2. Creates a folder called MaST-results-YEAR-DAY
3. Calculates quantile values for each test
4. Assigns points to schools based on 'Calc Quantile' column. Prints top 5 schools to screen and saves entire list to file in results folder.
5. Based on 'Award Quantile' column, creates excel files for each subject for using a mailmerge to print certificates for top student winners.
6. Creates a PDF where each page is for a school and list the scores for each test taken by participating students. 

#### Results mode notes

If the MaST-data-final-YEAR-DAY.xlsx data file changes for any reason, this script should be rerun. It will overwrite existing results files created for that day.

## MaST.py --help
```
usage: MaST.py [-h] {update,results} ...

MC Math & Science Tournament Grader & Analysis.

positional arguments:
  {update,results}  Specify either 'update' or 'results'.
    update          Update data and generate CSV file.
    results         Read in final csv file and tally results. No new data will be configured.

options:
  -h, --help        show this help message and exit
```

## MaST.py update --help
```
usage: MaST.py update [-h] [-a ASCII] [-k KEYS] [-i STUDENTS] [--no_lost]

options:
  -h, --help            			      show this help message and exit
  -a ASCII, --ascii ASCII			      Location of ascii file from Scantron. (Default: MaST-ascii.dat
  -k KEYS, --keys KEYS  			      Location of excel file with test keys. (Default: MaST-Keys.xlsx)
  -i STUDENTS, --students STUDENTS	Location of excel file with student registration information. (Default: MaST-Students.xlsx)
  --no_lost             			      Specify to not print missing tests. Helpful at beginning when not all scantron files have been processed.
```

## MaST.py results --help
```
usage: MaST.py results [-h] [-d DATA] [-i STUDENTS] [-s SCHOOLS]

options:
  -h, --help            			      show this help message and exit
  -d DATA, --data DATA  			      Location of final datat excel file made with update mode. (Default: MaST-Schools.xlsx)
  -i STUDENTS, --students STUDENTS	Location of excel file with student registration information. (Default: MaST-Students.xlsx)
  -s SCHOOLS, --schools SCHOOLS		  Location of excel file with school information. (Default: MaST-Schools.xlsx)
```
