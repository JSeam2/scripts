# -*- coding: utf-8 -*-
"""
UIST Analysis Overview Script
"""
import os
import openpyxl as xl
import csv
import pandas as pd
import re

# Key directories
maindir = 'C:/Users/ASUS/Documents/UIST/5. Processed Data/Touchpad_FLEX'
mainfile = maindir + '/FLEX_Recorded_Gestures_FINAL.xlsx'
outfile = maindir + '/FLEX_Recorded_Gestures_FINAL2.xlsx'
gesture = maindir + '/Gesture Cleaning'
csvdir = maindir + '/Pivot/CSV'      # we use csv


wb = xl.load_workbook(mainfile)
sheet_name = wb.get_sheet_names()[0]
sheet = wb.get_sheet_by_name(sheet_name)


# Deal with csv files first
"""
Average and STD DEV
Col C:N
"""
# list all csv files in the csv dir
csvfiles = os.listdir(csvdir)

#stand, walk, run
P1 = [sheet["A4"], 
      sheet["A23"], 
      sheet["A42"]]

# populate names
p_num = []

for cell in sheet["A4:A15"]:
    temp = cell[0].value
    temp = temp.split(' ')
    p_num.append(temp[0][1:])   


for mvt in range(0, len(P1)):
    cell_1 = P1[mvt]
    
    if mvt == 0:
        mvt_type = "Standing"
    elif mvt == 1:
        mvt_type = "Walking"
    elif mvt == 2:
        mvt_type = "Running"            
    
    for row in range(0, 12):
        # look for a ???_PIVOT.csv in PIVOT folder corresponding to p_num
        str_p_num = str(p_num[row])
        
        for file in csvfiles:
            match = re.search(r'^'+ str_p_num +'(\S{1,})'+ mvt_type +'_PIVOT.csv$', file)
            
            avg_vals = []
            stddev_vals = []
            
            # Open the file and find the avg and std dev values
            if match:
                print('\nReading file: '+match.group(0)+'\n')
                with open(csvdir + '/' + match.group(0), 'r') as f:
                    reader =  csv.reader(f)
                    
                    for line in reader:
                        if line[0].lower() == "average":
                            if line[5] != '':
                                avg_vals += line[1:6]
                            else:
                                avg_vals.append(line[1])
                                                 
                        
                        elif line[0].lower() == "standard deviation":
                            if line[5] != '':
                                stddev_vals += line[1:6]
                                
                            else:
                                stddev_vals.append(line[1])
                                
                        # write to workbook
                        if len(avg_vals) == 6 and len(stddev_vals) == 6:
                            print('Writing to workbook avg and std dev values\n')
                            print('------------')
                            avg_vals = list(map(float, avg_vals))
                            stddev_vals = list(map(float, stddev_vals))
                            all_vals = avg_vals + stddev_vals
                            
                            for col in range(2, 14):
                                cell_1.offset(row, col).value = all_vals[col-2]
                                            

"""
Ymax, Ymin, Xmax, Xmin
U:AN
"""

P3 = [sheet["U4"],
      sheet["U23"],
      sheet["U42"]]

for mvt in range(0, len(P3)):
    cell_1 = P3[mvt]
    
    if mvt == 0:
        mvt_type = "Standing"
    elif mvt == 1:
        mvt_type = "Walking"
    elif mvt == 2:
        mvt_type = "Running"
        
    for row in range(0, 12):
        # look for a ???_PIVOT.csv in PIVOT folder corresponding to p_num
        str_p_num = str(p_num[row])
        
        for file in csvfiles:
            match = re.search(r'^'+ str_p_num +'(\S{1,})'+ mvt_type +'_DATA.csv$', file)
            
            if match:
                print('\nReading file looking for Ymax, Ymin, Xmax, Xmin: '+match.group(0)+'\n')
                
                # in the following order
                # Ymax, Ymin, Xmax, Xmin
                # down left right tap up 
                           
                df1 = pd.read_csv(csvdir + '/' + match.group(0))
                
                df1 = df1.dropna(axis = 1, how = 'all') #Drop na on col
                df1 = df1.dropna(axis = 0, how = 'all') #Drop na on row
                df1 = df1[df1.iloc[:,1] == 'ok']
                
                # Filter by gesture
                df1_down = df1[df1.iloc[:,2] == "down"]
                df1_left = df1[df1.iloc[:,2] == "left"]
                df1_right = df1[df1.iloc[:,2] == "right"]
                df1_tap = df1[df1.iloc[:,2] == "tap"]
                df1_up = df1[df1.iloc[:,2] == "up"]
                
                # collect the values and then paste it into the xl sheet
                xy_val = []
                # y max
                down_ymax = df1_down.iloc[:,6].max()
                left_ymax = df1_left.iloc[:,6].max()
                right_ymax = df1_right.iloc[:,6].max()
                tap_ymax = df1_tap.iloc[:,6].max()
                up_ymax = df1_up.iloc[:,6].max()
                
                xy_val.append(down_ymax)
                xy_val.append(left_ymax)
                xy_val.append(right_ymax)
                xy_val.append(tap_ymax)
                xy_val.append(up_ymax)
                
                
                # y min
                down_ymin = df1_down.iloc[:,6].min()
                left_ymin = df1_left.iloc[:,6].min()
                right_ymin = df1_right.iloc[:,6].min()
                tap_ymin = df1_tap.iloc[:,6].min()
                up_ymin = df1_up.iloc[:,6].min()
                
                xy_val.append(down_ymin)
                xy_val.append(left_ymin)
                xy_val.append(right_ymin)
                xy_val.append(tap_ymin)
                xy_val.append(up_ymin)
                

                # x max
                down_xmax = df1_down.iloc[:,5].max()
                left_xmax = df1_left.iloc[:,5].max()
                right_xmax = df1_right.iloc[:,5].max()
                tap_xmax = df1_tap.iloc[:,5].max()
                up_xmax = df1_up.iloc[:,5].max()
                
                xy_val.append(down_xmax)
                xy_val.append(left_xmax)
                xy_val.append(right_xmax)
                xy_val.append(tap_xmax)
                xy_val.append(up_xmax)
                
                # x min
                down_xmin = df1_down.iloc[:,5].min()
                left_xmin = df1_left.iloc[:,5].min()
                right_xmin = df1_right.iloc[:,5].min()
                tap_xmin = df1_tap.iloc[:,5].min()
                up_xmin = df1_up.iloc[:,5].min()
                
                xy_val.append(down_xmin)
                xy_val.append(left_xmin)
                xy_val.append(right_xmin)
                xy_val.append(tap_xmin)
                xy_val.append(up_xmin)
                
                print('Writing to workbook Ymax, Ymin, Xmax, Xmin\n')
    
                            
                for col in range(0, 20):
                    cell_1.offset(row, col).value = xy_val[col]
                


"""
Percentage of OK
O:S
"""
print("+++++++++++++++++++++++++++++")
print("Reading gestures")
print("+++++++++++++++++++++++++++++")

P2 = [sheet["O4"], 
      sheet["O23"], 
      sheet["O42"]]

# list all xlsx files in the directory recursively
xlfiles = []
for root, directory, files in os.walk(gesture):
    path = root.split(os.sep)
    #print((len(path) - 1) * '---', os.path.basename(root))
    
    for file in files:
        #print(file[-5:])
        if file.endswith('.xlsx'):
            xlfiles.append(os.path.join(root,file))

# Replace \\ with / for consistency
xlfiles = [xlfile.replace('\\', '/') for xlfile in xlfiles]


for mvt in range(0, len(P2)):
    cell_1 = P2[mvt]
    
    if mvt == 0:
        mvt_type = "Standing"
    elif mvt == 1:
        mvt_type = "Walking"
    elif mvt == 2:
        mvt_type = "Running"

    for row in range(0, 12):
        # look for a ???(cleaned).xlsx in gesture folder corresponding to p_num
        str_p_num = str(p_num[row])
        
        for xlfile in xlfiles:
            # need to settle regex
            xlmatch = re.search(r'^((?:[^/]*/)*)(.*)' + str_p_num 
                                + '\S{1,}' + mvt_type + '.xlsx$'
                                , xlfile)
            
            if xlmatch:
                print('\nReading file looking for % in: '+xlmatch.group(0)+'\n')
                df2 = pd.read_excel(xlfile)
                ok_val = []
                
                df2 = df2.iloc[:,0:3]            
                df2 = df2.dropna(axis = 1, how = 'all') #Drop na on col
                df2 = df2.dropna(axis = 0, how = 'any') #Drop na on row
                
                # Total %
                c_total = df2["OK?"].value_counts()            
                assert len(c_total) == 2            
                gross_total = c_total[0] + c_total[1] #c[0] = ok, c[1] = not ok
                total_ok = c_total[0] / float(gross_total)
               
                 # Down %
                df2_down = df2[df2["Log"] == "down"]
                c_down = df2_down["OK?"].value_counts()  
                assert len(c_down) <= 2
                
                try:
                    down_total = c_down[0] + c_down[1]
                except IndexError:
                    down_total = c_down[0]
                
                down_ok = c_down[0] / float(down_total)
                
                # Left %
                df2_left = df2[df2["Log"] == "left"]
                c_left = df2_left["OK?"].value_counts()  
                assert len(c_left) <= 2
                
                try:
                    left_total = c_left[0] + c_left[1]
                except IndexError:
                    left_total = c_left[0]            
                
                left_ok = c_left[0] / float(left_total)
                
                # Right %
                df2_right = df2[df2["Log"] == "right"]
                c_right = df2_right["OK?"].value_counts()  
                assert len(c_right) <= 2
                
                try:
                    right_total = c_right[0] + c_right[1]
                except IndexError:
                    right_total = c_right[0]
                
                right_ok = c_right[0] / float(right_total)
                
                # Tap %
                df2_tap = df2[df2["Log"] == "tap"]
                c_tap = df2_tap["OK?"].value_counts()  
                assert len(c_tap) <= 2
                
                try:
                    tap_total = c_tap[0] + c_tap[1]
                except IndexError:
                    tap_total = c_tap[0]
                
                tap_ok = c_tap[0] / float(tap_total)
                
                # Up %
                df2_up = df2[df2["Log"] == "up"]
                c_up = df2_up["OK?"].value_counts()  
                assert len(c_up) <= 2
                
                try:
                    up_total = c_up[0] + c_up[1]
                except IndexError:
                    up_total = c_up[0]
                
                up_ok = c_up[0] / float(up_total)
                
                # just checking
                assert gross_total == (down_total
                                 + left_total
                                 + right_total
                                 + tap_total
                                 + up_total)
                
                ok_val.append(down_ok)
                ok_val.append(left_ok)
                ok_val.append(right_ok)
                ok_val.append(tap_ok)
                ok_val.append(up_ok)
                ok_val.append(total_ok)
                #print(ok_val)
                # input into cmain excel file in the following order
                # down left right tap up
                for col in range(0, 6):
                    cell_1.offset(row, col).value = ok_val[col]
         
            
# Save edits into a new workbook
print('--------------------------')
print('Saving new workbook in same folder')
print('--------------------------')
wb.save(outfile)    
print('Saved as ' + outfile)


    