import pandas as pd
import numpy as np
from tabulate import tabulate
from IPython.display import display, HTML
import os
import webbrowser

try:
    path = os.path.abspath('sample.html')
    url = 'file://' + path
except: 
    print('There is a problem with the file paths')

try:
    with open(path, 'w') as f:
        f.write("""<!DOCTYPE html>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
*{h2{margin-bottom: 0;
  margin-top:0;}
p{margin-top: 0;}}
*{
  margin:0.2;
}
.col-container {
  display: table;
  width: 100%;
  margin-top:0;
}
.col {
  display: table-cell;
  margin-top:0;
    padding: 10px;
}
table, th, td {
  border-style:solid;
  border-color: #96D4D4;
  border: 1px solid #96D4D4;
  border-collapse: collapse;}
</style>
</head>
<body>
    <font size="-8">""")

except:
    print('There was an error when writing to the html file')


try:
    #import data as Excel File
    xls = pd.ExcelFile("QLA.AH.Tril.F.Nov22.xlsx")
    #pdata = np.array(pd.read_excel(xls, 'Pupil Data'))

except:
    print('Could not read the Excel File')

# create np arrays of bio, chem and phys data
try:
    bio = np.array(pd.read_excel(xls, 'Biology'))
    chem = np.array(pd.read_excel(xls, 'Chemistry'))
    phys = np.array(pd.read_excel(xls, 'Physics'))
    t_scores = np.array(pd.read_excel(xls, 'Total Scores'))
except:
    print("""Could not read from the tabs: Biology, Chemistry, Physics or Total Scores""")

try:
    # create a np array with bio, chem and phys data
    subjects = np.array([bio, chem, phys])
except:
    print("""The input file could not be written as a numpy array.""") 
    
try:
    # count the number of pupils: k 
    k = 0
    list = subjects[0,10:,:]
    for j in list: 
        if isinstance(j[2], str) and j[2]!= 'Forename':
            k = k + 1
except:
    print('Could not count the number of pupils from tab: Biology')

#k = 3 # DELETE AFTER TEST SAMPLE

for pupil in range(k): # for each pupil

    try: 
        pname = str(list[pupil+1, 2])
        nameprint = str("""<h2>""" + list[pupil+1, 2] +""", """ )
    except:
        print("""Could not read the pupils's name""")
        nameprint = str("""<h2> Name unknown, """)
    try:
        classprint = str(list[pupil+1, 4] + """</h2>""")
    except:
        classprint = str("""class unknown </h2>""")
    try: 
        with open(path, 'a') as f:
            f.write(nameprint) # print their name
            f.write(classprint) # print their class
            f.write("""<div class="col-container">""")
    except: 
        print("Could not write to file")
              
    for sub in np.arange(3): # for bio, chem and phys#
        subject = np.array(['Biology', 'Chemistry', 'Physics'])[sub]
        #print(list[pupil+1, 2] + ': ' + subject)

        # Count the number of subsections 
        try:
            N = 10 # +10 for index where subsections start
            for i in subjects[sub,1,10:]: 
                if float(i) > 0: 
                    N = N + 1 
        except:
            print('Failed to count the number of subsections for ' + subject)
            
        # make an array of total scores and percentages for each paper, overall and their final grade
        #try: 
        scores = pd.array(subjects[sub,11+pupil,10:N])
        Null_check = pd.isnull(scores)
        if Null_check.sum() != 0:
            print('There is a missing value in the row for ' + list[pupil+1, 2] + "'s marks per question")
            for a in np.arange(scores.size):
                if Null_check[a] == True: 
                    scores[a] = 0
        if sub == 0:
            bio_score = np.array(scores, dtype = int)
            
        if sub == 1:
            chem_score = np.array(scores, dtype=int)

        if sub == 2:
            phys_score = np.array(scores, dtype = int)
        #except: 
        #    print('Could not read row for ' + list[pupil+1, 2] + "'s marks per question")

        # make arrays of the question number, AO type, description, maximum number of marks and pupil score
        try: 
            q_number = pd.array(subjects[sub,1,10:N])
            Null_check = pd.isnull(q_number)
            if Null_check.sum() != 0:
                print('There is an Null value in q_number')
                for a in np.arange(q_number.size):
                    if Null_check[a] == True: 
                        q_number[a] = ""
        except: 
            print('Could not read row for question number.')
        
        try: 
            AO_tag = pd.array(subjects[sub,2,10:N])
            Null_check = pd.isnull(AO_tag)
            if Null_check.sum() != 0:
                print('There is an Null value in the AO_tag')
                for a in np.arange(AO_tag.size):
                    if Null_check[a] == True: 
                        AO_tag[a] = ""
            AO_tag = np.array(AO_tag, dtype = str)
        except: 
            print('Could not read row for AO type.')
            
        try: 
            q_text = pd.array(subjects[sub,3,10:N])
            Null_check = pd.isnull(q_text)
            if Null_check.sum() != 0:
                print('There is an Null value in the q_text')
                for a in np.arange(q_text.size):
                    if Null_check[a] == True: 
                        q_text[a] = ""
            q_text = np.array(q_text, dtype = str)
        except: 
            print('Could not read row for question descriptions.')
            
        try: 
            q_max = pd.array(subjects[sub,4,10:N])
            Null_check = pd.isnull(q_max)
            if Null_check.sum() != 0:
                print('There is a missing value in the row for maximum marks per subsection.')
                for a in np.arange(q_max.size):
                    if Null_check[a] == True: 
                        q_max[a] = 6
            q_max = np.array(q_max, dtype=int)
        except: 
            print('Could not read row for maximum marks per subsection.')
            
               
        #set column headings
        headings = np.array([['','Question','Mark', 'Out of', 'Silly mistake', 'Need to revise']])
        if sub == 0:
            chem_score = np.array([])
            phys_score = np.array([])
        if sub ==1:
            phys_score = np.array([])
        scores_array = np.array([bio_score, chem_score, phys_score], dtype=object)
        scores_array = np.array(scores_array)
        #print(np.array(scores_array[sub]))
        #create and display the table
        table = np.transpose(np.array((q_number, q_text, scores_array[sub],q_max,np.full_like(q_number,''),np.full_like(q_number,'')),)) # array of non-heading table elements
        #string1 = str(tabulate(np.concatenate((headings, table)),tablefmt="html"))
        with open(path, 'a') as f:
            if sub == 0:
                f.write("""  <div class="col" style="background:white">
                <h3>Biology Paper 1</h3>""")
            else:
                if sub == 1:
                    f.write("""  <div class="col" style="background:white">
                    <h3>Chemistry Paper 1</h3>""")
                if sub == 2: 
                    f.write(""" <div class="col" style="background:white"> <h3>Physics Paper 1</h3>""")
            test = np.concatenate((headings, table)) # array of all table elements
            f.write("""<table> <tbody>""")
            for ii in np.arange(test.shape[0]): # for each row
                f.write("""<tr>""")
                if ii==0:
                    for iii in np.arange(test.shape[1]): # for each column of each row (cell)
                        f.write("""<td>""" + str(test[ii, iii]) + """</td>""")
                    f.write("""</tr>""")
                if ii==1:
                    for iii in np.arange(test.shape[1]-1): # for each column of each row (cell)
                        f.write("""<td>""" + str(test[ii, iii]) + """</td>""")
                    f.write("""<td rowspan=" """+str(N-9) +""" >""")
                    for i in np.arange(70):
                        f.write("""<br>""")
                    f.write(""" </td>""")
                    f.write("""</tr>""")
                if ii>1:
                    for iii in np.arange(test.shape[1]-1): # for each column of each row (cell)
                        if ii>0 and iii==2 and 100*test[ii,2]/test[ii,3]<30: 
                            f.write("""<td style="background-color:#FF0000"> """ + str(test[ii, iii]) + """</td>""")
                        elif ii>0 and iii==2 and 100*test[ii,2]/test[ii,3]>70:
                            f.write("""<td style="background-color:#00FF00"> """ + str(test[ii, iii]) + """</td>""")
                        elif ii>0 and iii==2:
                            f.write("""<td  style="background-color:#FFBF00">""" + str(test[ii, iii]) + """</td>""")
                        else:
                            f.write("""<td>""" + str(test[ii, iii]) + """</td>""")
                    f.write("""</tr>""")
            #for filltable in range(65-N):
            #    f.write("""<td>.</td><td></td><td> </td><td></td><td></td><td></td></tr>""")
            f.write("""</tbody>          </table></div>""")

    if sub == 2:
        with open(path, 'a') as f:
            f.write("""</div>

            <div class="col-container">
            """)

            for subj in np.arange(3): # for bio, chem and phys #
                subjectt = np.array(['Biology', 'Chemistry', 'Physics'])[subj]

                try: 
                    q_number = pd.array(subjects[subj,1,10:N])
                    Null_check = pd.isnull(q_number)
                    if Null_check.sum() != 0:
                        print('There is an Null value in q_number')
                        for a in np.arange(q_number.size):
                            if Null_check[a] == True: 
                                q_number[a] = ""
                except: 
                    print('Could not read row for question number.')
                
                try: 
                    AO_tag = pd.array(subjects[subj,2,10:N])
                    Null_check = pd.isnull(AO_tag)
                    if Null_check.sum() != 0:
                        print('There is an Null value in the AO_tag')
                        for a in np.arange(AO_tag.size):
                            if Null_check[a] == True: 
                                AO_tag[a] = ""
                    AO_tag = np.array(AO_tag, dtype = str)
                except: 
                    print('Could not read row for AO type.')
                    
                try: 
                    q_text = pd.array(subjects[subj,3,10:N])
                    Null_check = pd.isnull(q_text)
                    if Null_check.sum() != 0:
                        print('There is an Null value in the q_text')
                        for a in np.arange(q_text.size):
                            if Null_check[a] == True: 
                                q_text[a] = ""
                    q_text = np.array(q_text, dtype = str)
                except: 
                    print('Could not read row for question descriptions.')
                    
                try: 
                    q_max = pd.array(subjects[subj,4,10:N])
                    Null_check = pd.isnull(q_max)
                    if Null_check.sum() != 0:
                        print('There is a missing value in the row for maximum marks per subsection.')
                        for a in np.arange(q_max.size):
                            if Null_check[a] == True: 
                                q_max[a] = 6
                    q_max = np.array(q_max, dtype=int)
                except: 
                    print('Could not read row for maximum marks per subsection.')
                
                try: 
                    total_scores = pd.array(t_scores[7+pupil, 6:13])
                    Null_check = pd.isnull(total_scores)
                    if Null_check.sum() != 0:
                        print('There is a missing value in the row for ' + list[pupil+1, 2] + "'s total marks")
                        for a in np.arange(total_scores.size):
                            if Null_check[a] == True: 
                                total_scores[a] = 0
                    total_scores = np.array(total_scores)
                    
                except: 
                    print('Could not read row for total scores per subject.')
                
        
                AO1_tot = 0
                AO2_tot = 0
                AO3_tot = 0 
                RP_tot = 0 
                MS_tot = 0
                Xover_tot = 0
                AO1_pupil_percent = 0
                AO2_pupil_percent = 0
                AO3_pupil_percent = 0 
                RP_pupil_percent = 0 
                Xover_pupil_percent = 0

                                        
                try:
                    grade = str(t_scores[7+pupil, 13])
                except:
                    print("No grade listed for " + pname)
                    grade = "N/A"

                
                #try:
                for i in np.arange(AO_tag.size):           
                    if AO_tag[i]=='AO1':
                        AO1_tot +=  q_max[i]
                        AO1_pupil_percent += scores_array[subj][i]
                    else: 
                        if AO_tag[i] == 'AO2':
                            AO2_tot += q_max[i]
                            AO2_pupil_percent += scores_array[subj][i]
                        else: 
                            if AO_tag[i] == 'AO3':
                                AO3_tot += q_max[i]
                                AO3_pupil_percent += scores_array[subj][i]
                            else:
                                print("There is a missing AO tag in " + subjectt)

                    if 'RP:' in q_text[i]:
                        RP_tot += q_max[i]
                        RP_pupil_percent += scores_array[subj][i]
                #except:
                    #print('Could not read the marks from each AO and RP for ' + pname)
                    

                try:
                    inputstrings_XO = np.array(t_scores[6,15:18]) # make an array including the maximum Xover marks available
                    for d in np.arange(inputstrings_XO.size):
                        num = ""
                        for c in inputstrings_XO[d]:
                            if c.isdigit():
                                num = num + c
                        inputstrings_XO[d] = num # extract the maximum Xover marks available 
                    XO_max = np.array(inputstrings_XO, dtype=int) # create a np array of ints with the max Xover marks
                    
                    inputstrings_MS = np.array(t_scores[6,43:46]) # make an array including the maximum maths marks available
                    for d in np.arange(inputstrings_MS.size):
                        num = ""
                        for c in inputstrings_MS[d]:
                            if c.isdigit():
                                num = num + c
                        inputstrings_MS[d] = num # extract the maximum maths marks available 
                    inputstrings_MS = np.array(inputstrings_MS, dtype=int) # create a np array of ints with the max maths marks
                except:
                    print('Failed to read the data from the tab: Total Scores')


                try: 
                    AO1_pupil_percent = int(100 * AO1_pupil_percent / AO1_tot)
                except:
                    print("Could not calculate AO1 percentage for " + pname + "'s " + subjectt + " marks")
                    AO1_pupil_percent = int(0)

                try:
                    AO2_pupil_percent = int(100 * AO2_pupil_percent / AO2_tot)
                except:
                    print("Could not calculate AO2 percentage for " + pname + "'s " + subjectt + " marks")
                    AO2_pupil_percent = int(0)

                try:
                    AO3_pupil_percent = int(100 * AO3_pupil_percent / AO3_tot)
                except:
                    print("Could not calculate AO3 percentage for " + pname + "'s " + subjectt + " marks")
                    AO3_pupil_percent = int(0)

                try:
                    RP_pupil_percent = int(100 * RP_pupil_percent / RP_tot)
                except:
                    print("Could not calculate RP percentage for " + pname + "'s " + subjectt + " marks")
                    RP_pupil_percent = int(0)

                try:
                    XO_pupil_percent = int(100 * t_scores[7+pupil, 15+subj] / XO_max[subj])
                except:
                    print("Could not calculate Xover percentage for " + pname + "'s " + subjectt + " marks")
                    XO_pupil_percent = int(0)

                try:
                    MS_pupil_percent = int(100 * t_scores[7+pupil, 43+subj] / inputstrings_MS[subj])
                except:
                    print("Could not calculate Maths Skills percentage for " + pname + "'s " + subjectt + " marks")
                    MS_pupil_percent = int(0)

                try:
                    total_max = np.sum(q_max)
                except:
                    print("Failed to calculate the total marks available, assuming this paper is out of 70")
                    total_max = int(70) # if not calculable, assume they did Combined (70 marks available per paper)
                    
                try:    
                    total = np.sum(scores_array[subj])
                except:
                    print("Failed to calculate the pupil's total marks, assuming that " + pname + " got 0 marks.")
                    total = int(0) # if not calculable, assume they got 0 marks altogether

                test = np.array([['AO1', str(AO1_pupil_percent) + '''%'''],['AO2', str(AO2_pupil_percent) + '''%'''],['AO3', str(AO3_pupil_percent) + '''%'''],['Required Practicals', str(RP_pupil_percent) + '''%'''],['Maths Skills', str(MS_pupil_percent) + '''%'''],['Crossover', str(XO_pupil_percent) + '''%''']])
                
                f.write( """<div class="col" style="background:white"> <table> <tbody> <tr>""")
            
                for ii in np.arange(test.shape[0]):
                    for iii in np.arange(test.shape[1]):
                        f.write("""<td  style="background-color:#C4E1AF">""" + str(test[ii, iii]) + """</td>""")
                    f.write("""</tr>
                    """)
            
                f.write("""</tbody></table>

                """)

                test = np.array([['Total marks', str(total_scores[2*subj])],['Percentage', str(int(100*(total_scores[2*subj+1])))+'''%''']])
            
                f.write("""<table><tbody>""")
                for ii in np.arange(test.shape[0]):
                    f.write("""<tr>""")
                    for iii in np.arange(test.shape[1]):
                        f.write("""<td  style="background-color:#F1FC03">""" + str(test[ii, iii]) + """</td>""")
                    f.write("""</tr>
                    """)
                f.write("""</tbody> </table> </div>

                """)
                #f.write( string3)
                #f.write('''<br><br>''')

                if subj == 2:
                    f.write("""</div>""")

            string4 = str(tabulate(np.array([['Overall marks', str(total_scores[6])],['Overall Percentage', str(int(100*total_scores[6]/210))+'''%'''],['Overall Grade', str(grade)]]),tablefmt="html"))
            f.write('''<div> <center>'''+ string4+ '''</center>  </div> </body></html>''')