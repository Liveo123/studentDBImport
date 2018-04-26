import pandas as pd

import numpy as np
from pandas import ExcelWriter
from numpy import nan as Nan
import math

# First year of calendar as defined by PowerSchool
START_OF_CAL = 1982

# The first row in the Historical Grade in the template sheet to add stuff
TEMPL_START_ROW = 6

# Load up the student table
stud_xl = pd.ExcelFile("mycsv.xlsx")

df_courses = stud_xl.parse('Master Course List')
df_transcript = stud_xl.parse('Student Transcript')
df_grades = stud_xl.parse('Grade Table')

# Load up the Template worksheet - Historical Grades
hist_xl = pd.ExcelFile("Import.xls")

df_hist = hist_xl.parse("Historical Grades")

# Create a dictionary to lookup letter grades
let_to_GPA = {}
for index, row in df_grades.iterrows():
    let_to_GPA.update({row['Symbol'].strip(): row['Q Points']})

# Create a letter to percent dictionary TODO: Is this correct?
let_to_pcnt = {'A+': 100, 'A':97, 'A-':93, 'B+':90, 'B':87, 'B-':83, 'C+':80, 'C':77, 'C-':73, 'D+':70, 'D':67, 'D-':65, 'F':0, }

# Loop through rows, created each as we go
for row_cnt in range(0, 60):

    # Set up current row number from student transcript table
    curr_stud_row = TEMPL_START_ROW + row_cnt - 1
    sec_row_cnt = 0   # This is used for doubling up on rows when on 
    DoAnotherRow = True # This will get set to False if no other row needed.


    while DoAnotherRow == True:
        # Append a new blank row to the Data Frame
        s2 = pd.Series([Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan])
        df_hist = df_hist.append(s2, ignore_index = True)
        
        # Add Student Number
        df_hist['Student_Number'][curr_stud_row + sec_row_cnt] = df_transcript['Unique ID'][row_cnt]
        print(df_hist['Student_Number'][curr_stud_row + sec_row_cnt])        

        # Add Course Name
        df_hist['Course Name'][curr_stud_row + sec_row_cnt] = df_transcript['Course Name'][row_cnt]
        df_hist['Course Name'][curr_stud_row + sec_row_cnt].strip()[-2:]
        print(df_hist['Course Name'][curr_stud_row + sec_row_cnt]) 
        print("curr_stud_row = " + str(curr_stud_row))
        print("sec_row_cnt = " + str(sec_row_cnt))
        
        # Add Course Number
        df_hist['Course Number'][curr_stud_row + sec_row_cnt] = df_transcript['Course Number'][row_cnt]
        print (df_hist['Course Number'][curr_stud_row + sec_row_cnt]) 
        
        # EarnedCrHrs - 0.5, 1 TODO: How to calculate?
        df_hist['EarnedCrHrs'][curr_stud_row + sec_row_cnt] = 0.5 #df_transcript['Course Number'][row_cnt]
        print(df_hist['EarnedCrHrs'][curr_stud_row + sec_row_cnt]) 

        # Grade - A, B, C, D, F, NG
        # If this is the first row, use column 4 grade, otherwise use column 8
        # Also need to ensure not NaN
        if sec_row_cnt == 0:
            if str(df_transcript['RC Column 4'][row_cnt]).strip() != "nan":
                df_hist['Grade'][curr_stud_row + sec_row_cnt] = str(df_transcript['RC Column 4'][row_cnt]).strip()
        else: 
            if str(df_transcript['RC Column 8'][row_cnt]).strip() != "nan":
                df_hist['Grade'][curr_stud_row + sec_row_cnt] = str(df_transcript['RC Column 8'][row_cnt]).strip()
    
        # PotentialCrHrs - 0.5, 1 TODO: How to calculate?
        df_hist['PotentialCrHrs'][curr_stud_row + sec_row_cnt] = 0.5 #df_transcript['Course Number'][row_cnt]

        # Storecode - S1, T1, Y1, Q3 TODO:What is this?
        df_hist['Storecode'][curr_stud_row + sec_row_cnt] =  'S1' #df_transcript['Course Number'][row_cnt]

        # Termid - 1200 TODO: What is the term id
        curr_termid =  df_transcript['Calendar Year'][row_cnt] - START_OF_CAL 
        df_hist['Termid'][curr_stud_row + sec_row_cnt] = 1234

        # GPA Points - 4 TODO: What do we need to add with credit
        if pd.isnull(df_hist['Grade'][curr_stud_row + sec_row_cnt]) == False:
            df_hist['GPA Points'][curr_stud_row + sec_row_cnt] = let_to_GPA[df_hist['Grade'][curr_stud_row]]
        
        # Percent - 95 - Get the current letter grade from 'Grade' and
        # then convert to percent using dictionary let_to_pcnt
        if pd.isnull(df_hist['Grade'][curr_stud_row + sec_row_cnt]) == False:
            df_hist['Percent'][curr_stud_row + sec_row_cnt] = let_to_pcnt[df_hist['Grade'][curr_stud_row]]


        # SchoolName - GEMS American Academy
        df_hist['SchoolName'][curr_stud_row + sec_row_cnt] =  'GEMS American Academy'

        # Grade_Level - 10
        df_hist['Grade_Level'][curr_stud_row + sec_row_cnt] = df_transcript['Grade Level'][row_cnt]

        # Credit Type - Units, or MA
        df_hist['Credit Type'][curr_stud_row + sec_row_cnt] = 'Units' #df_transcript['Grade Level'][row_cnt]

        # Teacher Name - Mary Smith
        df_hist['Teacher Name'][curr_stud_row + sec_row_cnt] = df_transcript['Staff Name'][row_cnt]

        # Schoolid - 100 TODO: What is the school id?
        df_hist['Schoolid'][curr_stud_row + sec_row_cnt] =  5566 #df_transcript['Course Number'][row_cnt]

        # ExcludeFromGPA - 1 or 0 TODO: Is this always 0?  Same for two below
        df_hist['ExcludeFromGPA'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]
        
        # ExcludeFromClassRank - 1 or 0
        df_hist['ExcludeFromClassRank'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]
        
        # ExcludeFromHonorRoll - 1 or 0
        df_hist['ExcludeFromHonorRoll'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]

        # Need to check for courses that need two rows.  If this does, set 
        # Second Row Counter to 2 else
        blerh =  df_hist['Course Name'][curr_stud_row + sec_row_cnt].strip()[-2:]
        print('blerh = ' + blerh)
        if sec_row_cnt == 0 and blerh == 'HL':
            DoAnotherRow = True
            sec_row_cnt = 1

        else:
            DoAnotherRow = False

c = df_hist['Course Name'][curr_stud_row + sec_row_cnt]

# Write to a new Excel file
writer = ExcelWriter('NewFile.xlsx')
df_hist.to_excel(writer,'Sheet1',index=False)
writer.save()



