import pandas as pd
import numpy as np
import xlrd
from pandas import ExcelWriter
import PySimpleGUI as sg

""" NOTES
df = pd.read_html(HTMLfile)

- Only export "Space Load Summary" (with "Peak") from HAP.
- Save the HAP output as "Web Page, Filtered" so that it is only one file.

NOTES """

### Read HTML file ###
Input_ImportPath = sg.popup_get_file('HTML file to import:')
list_tables = pd.read_html(Input_ImportPath, flavor='html5lib')

# HTMLfile = r'C:\Users\kyleg\OneDrive - FFE Inc\Python\HVAC_HAP Load Output\Resources\Systems Design Report.htm'
# list_tables = pd.read_html(HTMLfile, flavor='html5lib')




### Function to get Space number & name. Returns a list. ###
def SpaceNames(tables):
    list_SpaceNames = []
    
    sub1 = 'Space'
    sub2 = ' In'

    for df in tables:
        titleString = df[0][0]
        
        idx1 = titleString.index(sub1)
        idx2 = titleString.index(sub2)

        SpaceName_str = ''
        # getting elements in between 'Space' & ' In'
        for idx in range(idx1 + len(sub1) + 1, idx2):
            SpaceName_str = SpaceName_str + titleString[idx]

        SpaceName_str = SpaceName_str.strip('"') # Remove quotation marks
        SpaceName_str = SpaceName_str.replace('  ', ' ') # Replace double space with single space

        list_SpaceNames.append(SpaceName_str)


    return list_SpaceNames


### Get list of Spaces from function ###
list_SpaceNames = SpaceNames(list_tables[::2])




### Combine list of "Component Load" DataFrames with Spaces as keys ###
df_ComponentLoads = pd.concat(list_tables[::2])

df_ComponentLoads.drop(range(1,23,1), axis=0, inplace=True) # Drop unwanted rows
df_ComponentLoads.drop([1,4,6], axis=1, inplace=True) # Drop unwanted columns


### Get only "Total Zone Loads" row and reset the index
df_TotalZoneLoads = df_ComponentLoads.loc[23].reset_index(drop=True)





### Create DataFrame for final output ###
Columns_SystemLoads = ['Space', 'Zone', 'Cooling Load_Sensible', 'Cooling Load_Latent', 'Heating Load'] # Columns for "df_HAP_Output"
df_HAP_Output = pd.DataFrame(columns=Columns_SystemLoads)




### Assign values to columns for final output ###
df_HAP_Output['Space'] = list_SpaceNames
df_HAP_Output['Cooling Load_Sensible'] = df_TotalZoneLoads[2]
df_HAP_Output['Cooling Load_Latent'] = df_TotalZoneLoads[3]
df_HAP_Output['Heating Load'] = df_TotalZoneLoads[5]




### Get file path and sheet name from user input. ###
Input_excelpath = sg.popup_get_file('Excel file to write to:')
Input_sheetname = sg.popup_get_text('Excel Sheet Name:', default_text='PythonImport')


### Write data to excel file. ###
with ExcelWriter(Input_excelpath, mode="a", if_sheet_exists='replace') as writer:
    df_HAP_Output.to_excel(writer, sheet_name=Input_sheetname)
