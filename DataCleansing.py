import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

file = pd.ExcelFile('C:/Users/Gaurav/Desktop/attachment_attachment_Numpy_Pandas_Assignment_updated/Numpy_Pandas_Assignment/obes-phys-acti-diet.xls')
print (file.sheet_names)        # Fetching number of Sheets present in Excel file

## Table 7.1 Finished Admission Episodes with a primary diagnosis of obesity, by gender, 2000/01 to 2011/12		
x = file.parse(u'7.1', skiprows=5, skip_footer=14, names=['Year','Total','Males','Females','Nan', 'None'])   # Parsing sheet 7.1
#print (x)
drop_7_1 = x.dropna(axis=1).set_index('Year')        # Drop NaN values
print (drop_7_1)
print (drop_7_1.describe())                          # Fetching mean, std, etc

## Table 7.2 Finished Admission Episodes with a primary diagnosis of obesity, by age group, 2000/01 to 2011/12									
y =  file.parse(u'7.2', skiprows=4, skipfooter=14)
#print (y)
index_rename = y.rename(columns={'Unnamed: 0': 'Year'})     # Switching Index from by-default to Year
#print (index_rename)
cleaned_7_2 = index_rename.drop(index_rename.index[0])      # Drop Row using indice
drop_na = cleaned_7_2.dropna(axis=1).set_index('Year')      # Drop NaN values Column Wise
print (drop_na)
print (drop_na.describe())

## Table 7.4 Finished Admission Episodes with a primary or secondary diagnosis of obesity, by gender, 2000/01 to 2011/12			
z = file.parse(u'7.4', skiprows=3, skipfooter=14)
#print (z)
index_rename1 = z.rename(columns={'Unnamed: 0': 'Year'})
#print (index_rename1)
cleaned_7_4 = index_rename1.drop(index_rename1.index[0])
print (cleaned_7_4)
drop_na1 = cleaned_7_4.dropna(axis=1).set_index('Year')
print (drop_na1)
print (drop_na1.describe())

## Table 7.5 Finished Admission Episodes with a primary or secondary diagnosis of obesity, by age group, 2000/01 to 2011/12									
a = file.parse(u'7.5', skiprows=3, skipfooter=14)
#print (a)
index_rename2 = a.rename(columns={'Unnamed: 0':'Year'})
#print (index_rename2)
cleaned_7_5 = index_rename2.drop(index_rename2.index[0])
#print (cleaned_7_5)
drop_na2 = cleaned_7_5.dropna(axis=1).set_index('Year')
print (drop_na2)
print (drop_na2.describe())

## Table 7.7 Finished Consultant Episodes with a primary diagnosis of obesity and a main or secondary procedure of 'Bariatric Surgery' by gender, 2000/01 to 2011/12
b = file.parse(u'7.7', skiprows=3, skipfooter=18)
#print (b)
drop_row = b.drop(b.index[0:1])
#print (drop_row)
index_rename3 = drop_row.rename(columns={'Unnamed: 0': 'Year'})
#print (index_rename3)
del index_rename3['Unnamed: 4']
del index_rename3['Unnamed: 5']
des = index_rename3.head(None)
#print (des)
cleaned_7_7 = des.drop(des.index[7]).set_index('Year')
print (cleaned_7_7)
#fillna_7_7 = cleaned_7_7.fillna({'Total':0, 'Male':0, 'Female':0})
print (cleaned_7_7.describe())

# CODED BY - GAURAV_PADAWE
