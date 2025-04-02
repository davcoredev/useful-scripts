#! files_merge_script.py

"""
This script reads all the Excel files and merges them into a single file.

--- IMPORTANT, All files must have the same scheme/structure ---

The desired output can be changed using the DESIRED_COLUMNS variable and also by modifying the sheet_name='' parameter.
"""

from datetime import datetime
import pandas as pd
import logging
import os

# Added logging configuration to replace prints statements
logging.basicConfig(filename=f'logs/{datetime.today().strftime("%H_%M_%S_%d_%m_%Y")}-files_merge_script.log', 
                    level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')

#TODO: Insert path here, eg. "C:/User/Documents/Files" 

PATH_TO_FILES = ""

#TODO: Insert the columns to keep from the files
DESIRED_COLUMNS = [
    "Column 1",
    "Column 2",
    "Column 3",
] 

def get_files_in_path(path:str)->list[str]:
    """
    Function to extract the name and path for files in a folder and subfolders
    """
    # Read all the files in the directory that matches the .endswith parameter, in this case Excel files (.xlsm)
    excel_files = [os.path.join(root, name) 
                for root, dirs, files in os.walk(path) 
                for name in files if name.endswith('.xlsm')]
    
    # TODO: Remove "PLACEHOLDER" with the needed word, this can be usefull if we only want to read files starting with an specific word, eg. "sales-july-2024.xlsm"
    # Remove the files that doesn't start with "PLACEHOLDER" and contains '~' in its name (temporary files)
    cleaned_excel_files = [file_name for file_name in excel_files 
                           if "PLACEHOLDER" in file_name and '~' not in file_name]
    
    # Save into the log the number of files found
    logging.info(f'Found {len(cleaned_excel_files)} files.')

    return cleaned_excel_files


def get_files_content(files_path:list)->list:
    """
    Function to read the files found in the get_files_in_path(path:str) function.
    """
    dataframes = []
    
    for i, file in enumerate(files_path, start=1):
        try:
            # Read the details from the sheet 'IND' and aditional details from the sheet 'Reg'
            # TODO: replace "SHEET_NAME1" and "SHEET_NAME2" with real sheet names
            # the next line is used to read some specific sections of the book and ad them later to the dataframe
            # specific_details = pd.read_excel(file, sheet_name='SHEET_NAME1', engine='calamine').iloc[[0,1],6]
            df = pd.read_excel(file, sheet_name='SHEET_NAME2', skiprows=8, nrows=50)
            
            # Adds two columns to the Dataframe using the data from line 60
            # TODO: replace "SPECIFIC_DETAIL1" and "SPECIFIC_DETAIL2" with the name of the actual columns
            # df['SPECIFIC_DETAIL1'] = specific_details.iloc[1]
            # df['SPECIFIC_DETAIL2'] = specific_details.iloc[0]
            
            # Adds the Dataframe to the list
            dataframes.append(df)
            
            # Add the number and name of the file to the logs
            logging.info(f'Processed file {i}: {file}')
        except Exception as e:
            # In case of error, write to the logs
            logging.error(f'Error processing file {file}: {e}')
    return dataframes

def initiate()->None:
    """
    Initialize the script:
        1 - Extract the details of the files (name and path) in the selected directory
        2 - Reads all the files and saves them into a list 
    """
    _files = get_files_in_path(PATH_TO_FILES)
    df_list = get_files_content(_files)
    
    if df_list:
        # Get all the dataframes and merg them into one, drop the duplicates and then saves the final dataframe into the desired directory
        final_df = pd.concat(df_list, ignore_index=True)
        final_df = final_df[DESIRED_COLUMNS].drop_duplicates()
        # TODO: add the actual path and dame of the file to be saved
        final_df.to_excel("", index=False)
        
        # If success, write to the logs
        logging.info(f'Done {len(_files)} files processed.')
    else:
        # In case of error, write to the logs
        logging.warning('No dataframes to concatenate')

if __name__ == '__main__':
    initiate()    