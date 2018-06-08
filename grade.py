import pandas as pd

import numpy as np
from pandas import ExcelWriter
from numpy import nan as Nan
import math
import xlrd
import sys
import time
import datetime

timing = []
t0 = []
t1 = []
t2 = []
t3 = []
t4 = []
t5 = []

BIG_TIME = time.time()

def saveHist():
    writer = ExcelWriter('NewFile.xlsx')
    df_hist.to_excel(writer,'Sheet1',index=False)
    writer.save()

# Are we running this for the first or 2nd semester (1 or 2)
S1_OR_S2 = 2

# First year of calendar as defined by PowerSchool
START_OF_CAL = 1990

# The first row in the Historical Grade in the template sheet to add stuff
TEMPL_START_ROW = 5

# ID of High School
SCHOOL_ID = 3

# Course lengths
LENGTH_ALL = 'ALL'
LENGTH_SEM = 'SEM'

# Number of columns for various fields in their respective tables
LENGTH_FIELD_NO = 13

# Higher Level and Standard Level course name post-fixes
HIGHER_LEVEL = 'HL'
STANDARD_LEVEL = 'SL'

### The logic switches ###
SEC_ROW = False
TWO_ROWS = False

# Load up the student table
stud_xl = pd.ExcelFile("mycsv23.xlsx")

df_courses = stud_xl.parse('Master Course List')
df_transcript = stud_xl.parse('Student Transcript')
df_grades = stud_xl.parse('Grade Table')
df_student = stud_xl.parse('Student') 

# Load up the Template worksheet - Historical Grades
hist_xl = pd.ExcelFile("Import.xlsx")

df_hist = hist_xl.parse("Historical Grades")

# Create a dictionary to lookup letter grades
let_to_GPA = {}
for index, row in df_grades.iterrows():
    let_to_GPA.update({row['Symbol'].strip(): row['Q Points']})

# Create a letter to percent dictionary
# TODO: What grade for P?
let_to_pcnt = {'A+': 100, 'A':96, 'A-':93, 'B+':89, 'B':86, 'B-':83, 'C+':79, 'C':75, 'C-':70, 'D+':65, 'D':60, 'D-':55, 'F':49, 'P':0}

# TODO: What grade for P?
# Create a Letter to GPA dictionary
let_to_gpa = {'A+': 4.30, 'A':4.00, 'A-':3.70, 'B+':3.30, 'B':3.00, 'B-':2.70, 'C+':2.30, 'C':2.00, 'C-':1.70, 'D+':1.30, 'D':1.00, 'D-':0.70, 'F':0.00, 'P':0}

# Incremented each time there is a double row student (e.g. for HL)
#extra_row_cnt = 0

# Need a count similar to act_row_cnt, but that gets incremented when a new row 
# gets created in the new file/dataframe
newfile_row_cnt = -1


act_row_cnt = -1
curr_stud_row = TEMPL_START_ROW

# Loop through rows, creating each as we go
for row_cnt in range(0, 55):

    act_row_cnt += 1
    
    # Don't bother with anything if there is are no grades for the student.  Go 
    # to the next row.
    acn = df_transcript['Course Name'][act_row_cnt]
    ac4 = df_transcript['RC Column 4'][act_row_cnt]
    if (S1_OR_S2 == 1 and 
       pd.isnull(df_transcript['RC Column 4'][act_row_cnt]) == False) or \
       (S1_OR_S2 == 2 and \
       pd.isnull(df_transcript['RC Column 8'][act_row_cnt]) == False):

        # Increment no. of rows in new file / Dataframe
        newfile_row_cnt += 1
        
        # Create/update main row counter for spreadsheet row where current 
        # student will be placed
        curr_stud_row = newfile_row_cnt + TEMPL_START_ROW

        # Temp
        crs_num = df_transcript['Course Number'][curr_stud_row]

##### START MAIN SECTION #####

        # Append a new blank row to the Data Frame
        s2 = pd.Series([Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan,Nan])
        df_hist = df_hist.append(s2, ignore_index = True)
        
        # Testing - Writing to disk - REMOVE LATER
        # Add Student Number
        df_hist['Student_Number'][curr_stud_row ] = df_transcript['Unique ID'][act_row_cnt]

        # Add Course Number
        crs_number_temp = df_transcript['Course Number'][act_row_cnt]
        mycnt = len(df_student.loc[df_student['UNIQUE ID'] == df_transcript['Unique ID'][act_row_cnt]])
        df_hist['Course Number'][curr_stud_row] = df_student.loc[df_student['UNIQUE ID'] == df_transcript['Unique ID'][act_row_cnt]].iloc[0]['Bluebook ID']

        # Add Course Name
        tempCN = df_transcript['Course Name'][act_row_cnt]
        df_hist['Course Name'][curr_stud_row] = str(df_courses.loc[df_courses['Course Number'] == df_transcript['Course Number'][act_row_cnt]].iloc[0]['Description'])
        
        # Termid - Year of work - start of PowerSchool time * 1000
        cal_year = int(df_transcript['Calendar Year'][act_row_cnt])
        curr_termid = (cal_year - START_OF_CAL) * 100
        df_hist['Termid'][curr_stud_row] = str(curr_termid)
        
        # SchoolName - GEMS American Academy
        df_hist['SchoolName'][curr_stud_row] =  'GEMS American Academy'

        # TODO: grade level = grade level - (current year - caledndar year of course)
        # Grade_Level - 10
        grade_lvl = df_transcript['Grade Level'][act_row_cnt]
        
        # TODO: Remove G on grade level if one has been added.
        #if str(grade_lvl[0]) == 'G':
            #grade_lvl = grade_lvl[1:]

        df_hist['Grade_Level'][curr_stud_row] = int(grade_lvl) - int(df_transcript['Relative Year'][act_row_cnt])
        #df_hist['Grade_Level'][curr_stud_row] = int(grade_lvl) \
                                                #+ int(datetime.datetime.now().year) \
                                                #- cal_year

        #NOT NEEDED BY NEW SYSTEM
        # Credit Type - Units, or MA
        #df_hist['Credit Type'][curr_stud_row] = 'Units' #df_transcript['Grade Level'][act_row_cnt]

        # Teacher Name - Mary Smith
        df_hist['Teacher Name'][curr_stud_row] = df_transcript['Staff Name'][act_row_cnt]

        # Schoolid
        df_hist['Schoolid'][curr_stud_row] =  SCHOOL_ID 

        # ExcludeFromGPA - 1 or 0
        df_hist['ExcludeFromGPA'][curr_stud_row] = 0 #df_transcript['Course Number'][act_row_cnt]
        
        # ExcludeFromClassRank - 1 or 0
        df_hist['ExcludeFromClassRank'][curr_stud_row] = 0 #df_transcript['Course Number'][act_row_cnt]
        
        # ExcludeFromHonorRoll - 1 or 0
        df_hist['ExcludeFromHonorRoll'][curr_stud_row] = 0 #df_transcript['Course Number'][act_row_cnt]


##### END MAIN SECTION #####
##### START LOGIC SECTION #####
        
        # Are there going to be 2 rows?  This occurs if there are two grades.
        #print("col 4 = {}".format(df_transcript['RC Column 4'][act_row_cnt])) 
        #print("col 8 = {}".format(df_transcript['RC Column 8'][act_row_cnt])) 
        #if pd.isnull(df_transcript['RC Column 4'][act_row_cnt]) == False and \
           #pd.isnull(df_transcript['RC Column 8'][act_row_cnt]) == False:
                #TWO_ROWS = True
        #else:
                #TWO_ROWS = False

        #### Output this stuff


        # Do we start on the first or second grade?
        #if pd.isnull(df_transcript['RC Column 4'][act_row_cnt]) == False:
            #SEC_ROW = False
        #elif pd.isnull(df_transcript['RC Column 8'][act_row_cnt]) == False:
            #SEC_ROW = True


##### END LOGIC SECTION #####
##### START EXTRA SECTION #####

        # Loop around one or two rows, depending on the grades.  There can be either both grades,
        # (SEC_ROW == False and TWO_ROWS = True) a grade for first semester (SEC_ROW == False 
        # and TWO_ROWS = False) or a grade for second semester (SEC_ROW == True and TWO_ROWS = True)
        
        
        #complete = False


        ## Start Timer
        #tmr = time.time()
   
        # Storecode - S1, S2
        if (S1_OR_S2 == 1):
            df_hist['Storecode'][curr_stud_row] =  'S1'
        else:
            df_hist['Storecode'][curr_stud_row] =  'S2'


        # Grade - A, B, C, D, F, NG
        # If this is the first row, use column 4 grade, otherwise use column 8
        # Also need to ensure not NaN
        #if SEC_ROW == False:
            #if pd.isnull(df_transcript['RC Column 4'][act_row_cnt]) == False:
        if S1_OR_S2 == 1:
                df_hist['Grade'][curr_stud_row] = str(df_transcript['RC Column 4'][act_row_cnt]).strip()
        else: 
            #if pd.isnull(df_transcript['RC Column 8'][act_row_cnt]) == False:
            df_hist['Grade'][curr_stud_row] = str(df_transcript['RC Column 8'][act_row_cnt]).strip()

        
        # Percent - 95 - Get the current letter grade from 'Grade' and
        # then convert to percent using dictionary let_to_pcnt
        if pd.isnull(df_hist['Grade'][curr_stud_row]) == False:
            df_hist['Percent'][curr_stud_row] = let_to_pcnt[df_hist['Grade'][curr_stud_row]]
        else:
            print("Empty grade for {}, course {}".format(df_hist['Student_Number'][curr_stud_row], df_hist['Course Number'][curr_stud_row]))
            #sys.exit()

        # PotentialCrHrs - 0.5, 1 Is this logic correct - double check for HL students with only one semester
        # To find the potential credit, need to find the relevant row in the master course list (where the name
        # of this course is the same) and find the CRDTS from there.  Also need to find the length (SEM or ALL).
        # If it is SEM, potential credits is 1*CRDTS, else if length is FULL, credits is 0.5*CRDTS
        # This gives errors if > course with the same name e.g. SocDP1ToK 
        # So, first find row in df_courses where the course name is the same as this one.

        course_crdts = float(df_courses.loc[df_courses['Course Number'] == df_transcript['Course Number'][act_row_cnt]].iloc[0]['CRDTS'])
        course_length = str(df_courses.loc[df_courses['Course Number'] == df_transcript['Course Number'][act_row_cnt]].iloc[0]['Length'])

        if str(course_length) == str('SEM'): #LENGTH_SEM:
            df_hist['PotentialCrHrs'][curr_stud_row] = course_crdts
        elif str(course_length) == str('ALL'): #LENGTH_ALL:
            df_hist['PotentialCrHrs'][curr_stud_row] = 0.5 * course_crdts
        elif str(course_length) == str('QTR'): #TODO Should we be including quarters?
            df_hist['PotentialCrHrs'][curr_stud_row] = 0.25 * course_crdts
                    
        # If grade was a pass, earned credit is the potential credit, otherwise 0.            
        # EarnedCrHrs - 0.5, 1 Problems with GPA again
        if pd.isnull(df_hist['Grade'][curr_stud_row]) == False and df_hist['Grade'][curr_stud_row] != 'F':
            print(df_hist['PotentialCrHrs'][curr_stud_row])
            df_hist['EarnedCrHrs'][curr_stud_row] = df_hist['PotentialCrHrs'][curr_stud_row]
                #df_courses.loc[df_courses['Name'] == df_transcript['Course Name'][act_row_cnt], 'CRDTS'].iloc[0]


        # GPA Points - 4 

        gpa_temp = 0.0
        if str(df_hist['Course Name'][curr_stud_row])[-2:] == HIGHER_LEVEL:
            gpa_temp = 0.5
        elif str(df_hist['Course Name'][curr_stud_row])[-2:] == STANDARD_LEVEL:
            gpa_temp = 0.25
        else:
            gpa_temp = 0.0

        # Use the grade the student received to find the points from the grade table and 
        # then add them to the earned grade
        gpa_temp += let_to_gpa[df_hist['Grade'][curr_stud_row]]
        df_hist['GPA Points'][curr_stud_row] = gpa_temp
        # Do we need to loop again?
        if (SEC_ROW == True) or (SEC_ROW == False and TWO_ROWS == False):
            # If two rows were included, add extra to current row count.
            if SEC_ROW and TWO_ROWS:
                extra_row_cnt += 1
            complete = True
        else:
            # Get ready for second row.
            if SEC_ROW == False and TWO_ROWS == True:
                SEC_ROW = True

        
##### END EXTRA SECTION #####

# Write to a new Excel file
writer = ExcelWriter('NewFile4.xlsx')
df_hist.to_excel(writer,'Sheet1',index=False)
writer.save()
