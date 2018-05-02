import pandas as pd

import numpy as np
from pandas import ExcelWriter
from numpy import nan as Nan
import math
import xlrd

# First year of calendar as defined by PowerSchool
START_OF_CAL = 1990

# The first row in the Historical Grade in the template sheet to add stuff
TEMPL_START_ROW = 6

# ID of High School
SCHOOL_ID = 3

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

# Create a letter to percent dictionary 
let_to_pcnt = {'A+': 100, 'A':96, 'A-':93, 'B+':89, 'B':86, 'B-':83, 'C+':79, 'C':75, 'C-':70, 'D+':65, 'D':60, 'D-':55, 'F':49, }

# Incremented each time there is a double row student (e.g. for HL)
extra_row_cnt = 0

# Loop through rows, creating each as we go
for row_cnt in range(0, 60):

    # Set up current row number from student transcript table 
    # TEMPL_START_ROW - Size of initial rows in the template that we should ignore
    # row_cnt - current row in other tables (from loop)
    # extra_row_cnt - Incremented each time there is a double row student (e.g. for HL)
    curr_stud_row = TEMPL_START_ROW + row_cnt + extra_row_cnt - 1

    sec_row_cnt = 0   # This is used for doubling up on rows when on 
    DoAnotherRow = True # This will get set to False if no other row needed.


    while DoAnotherRow == True:

        # Is the student a IB Diploma Higher Level student?
        if df_transcript['Course Name'][curr_stud_row + sec_row_cnt].strip()[-2:] == 'HL':
            isHL = True
        else:
            isHL = False

        # Append a new blank row to the Data Frame
        s2 = pd.Series([Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan])
        df_hist = df_hist.append(s2, ignore_index = True)
        
        # Add Student Number
        df_hist['Student_Number'][curr_stud_row + sec_row_cnt] = df_transcript['Unique ID'][row_cnt]

        # Add Course Name
        df_hist['Course Name'][curr_stud_row + sec_row_cnt] = df_transcript['Course Name'][row_cnt]
        #df_hist['Course Name'][curr_stud_row + sec_row_cnt].strip()[-2:]
        #print(df_hist['Course Name'][curr_stud_row + sec_row_cnt]) 
        #print("curr_stud_row = " + str(curr_stud_row))
        #print("sec_row_cnt = " + str(sec_row_cnt))
        
        # Add Course Number
        df_hist['Course Number'][curr_stud_row + sec_row_cnt] = df_transcript['Course Number'][row_cnt]
        
        # EarnedCrHrs - 0.5, 1 TODO: Problems with GPA again
        if pd.isnull(df_hist['Grade'][curr_stud_row + sec_row_cnt]) and df_hist['Grade'][curr_stud_row + sec_row_cnt] != 'F':
            #df.loc[df['B'] == 3, 'A']
            ff = (df_courses.loc[df_courses['Name'] == df_transcript['Course Name'][row_cnt], 'CRDTS'])
            print("hhhhhhhhhhhhhh")
            print(ff)
            print("hhhhhhhhhhhhhh")
            #df_hist['EarnedCrHrs'][curr_stud_row + sec_row_cnt] =  \
                #df_courses.loc[df_courses['Name'] == df_transcript['Course Name'][row_cnt], 'CRDTS'].iloc[0]

        # Grade - A, B, C, D, F, NG
        # If this is the first row, use column 4 grade, otherwise use column 8
        # Also need to ensure not NaN
        if sec_row_cnt == 0:
            if str(df_transcript['RC Column 4'][row_cnt]).strip() != "nan":
                df_hist['Grade'][curr_stud_row + sec_row_cnt] = str(df_transcript['RC Column 4'][row_cnt]).strip()
        else: 
            if str(df_transcript['RC Column 8'][row_cnt]).strip() != "nan":
                df_hist['Grade'][curr_stud_row + sec_row_cnt] = str(df_transcript['RC Column 8'][row_cnt]).strip()
    
        # PotentialCrHrs - 0.5, 1 TODO: Is this logic correct - double check for HL students with only one semester
        if pd.isnull(df_hist['Grade'][curr_stud_row + sec_row_cnt]) and df_hist['Grade'][curr_stud_row + sec_row_cnt] != 'F':
            df_hist['PotentialCrHrs'][curr_stud_row + sec_row_cnt] = 0.5
        else:
            df_hist['PotentialCrHrs'][curr_stud_row + sec_row_cnt] = 0.0

        # Storecode - S1, T1, Y1, Q3 TODO: Is this logic correct - double check for HL students with only one semester
        if (sec_row_cnt == 0):
            df_hist['Storecode'][curr_stud_row + sec_row_cnt] =  'S1'
        else:
            df_hist['Storecode'][curr_stud_row + sec_row_cnt] =  'S2'

        # Termid - Year of work - start of PowerSchool time * 1000
        cal_year = int(df_transcript['Calendar Year'][row_cnt])
        curr_termid = cal_year - START_OF_CAL * 100
        df_hist['Termid'][curr_stud_row + sec_row_cnt] = str(curr_termid)

        # GPA Points - 4 
        # if not HL student, calculate by semester 1 grade + 
        #if pd.isnull(df_hist['Grade'][curr_stud_row + sec_row_cnt]) == False:
        #    if isHL = False:
        #        df_hist['GPA Points'][curr_stud_row + sec_row_cnt] = let_to_GPA[df_hist['Grade'][curr_stud_row]]

        
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

        # Schoolid
        df_hist['Schoolid'][curr_stud_row + sec_row_cnt] =  SCHOOL_ID #df_transcript['Course Number'][row_cnt]

        # ExcludeFromGPA - 1 or 0
        df_hist['ExcludeFromGPA'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]
        
        # ExcludeFromClassRank - 1 or 0
        df_hist['ExcludeFromClassRank'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]
        
        # ExcludeFromHonorRoll - 1 or 0
        df_hist['ExcludeFromHonorRoll'][curr_stud_row + sec_row_cnt] = 0 #df_transcript['Course Number'][row_cnt]

        # Need to check for courses that need two rows.  If this does, set 
        # Second Row Counter to 2 else


        #TODO: Check this logic is correct...
        # Check if a student needs a second row.  This would occur if a student is a IB Diploma Higher Level and
        # there is a grade in the second column (and this loop isn't already looking at the second column)
        if sec_row_cnt == 0 and isHL and pd.isnull(df_transcript['RC Column 4'][row_cnt]) == False:
            DoAnotherRow = True
            sec_row_cnt = 1
        else:
            DoAnotherRow = False
            if(sec_row_cnt == 1):
                extra_row_cnt += 1      # Used when extra row added for students that need them e.g. HL
    
c = df_hist['Course Name'][curr_stud_row + sec_row_cnt]

# Write to a new Excel file
writer = ExcelWriter('NewFile.xlsx')
df_hist.to_excel(writer,'Sheet1',index=False)
writer.save()



