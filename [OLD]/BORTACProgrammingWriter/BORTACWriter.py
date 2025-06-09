"""
Author: Hannah Layton
Last Updated: 05/12/2025 by Hannah Layton
"""
import subprocess
import sys

def install(package):
    """
    Automatically runs the Python pip install command to download necessary external packages

    Parameter: package (str) - package to install
    """
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# If user does not have any packages installed, install them automatically
try:
    from tkinter import Tk
    import os
    import tkinter.filedialog
    import tkinter.simpledialog
    import pandas as pd
    import numpy as np
    import re
    import openpyxl
except:
    install("pandas")
    install("numpy")
    install("openpyxl")

def askForTextInput(message):
    """
    Opens a window to enter and submit text

    Parameter:
    message (str) - prompt for user

    Return:
    (str) - user-inputted text
    """
    Tk().withdraw()
    user_input = tkinter.simpledialog.askstring(title="", prompt=message)
    return user_input

def getSCADAMap():
    """
    Opens a window to select a SCADA Map

    Return:
    (DataFrame) - SCADA Map read in as a DataFrame
    """
    Tk().withdraw()
    scadamappath = tkinter.filedialog.askopenfilename(title="Please select your SCADA Map.", filetypes=[("Excel files", "*.xlsx *.xls")])
    return pd.read_excel(scadamappath, sheet_name='BINARY OUTPUTS')

def cleanSCADAMap(scadamapdf):
    """
    Performs SCADA Map data cleaning

    Parameter:
    scadamapdf (DataFrame) - unclean SCADA Map read in as a DataFrame

    Return:
    (DataFrame) - SCADA Map as a cleaned DataFrame
    """
    # Drop entries with no wordbit
    wordbitcolumn = askForTextInput("Enter the column header for Wordbits in the SCADA Map: ")
    scadamapdf = scadamapdf.dropna(subset=[wordbitcolumn])
    scadamapdf = scadamapdf.dropna(axis=1, how='all')
    scadamapdf = scadamapdf.reset_index(drop=True) # Reset index of DataFrame for easier referencing in the future

    # The Slave IED Device listed in the SCADA Map can be converted into a matching format to Data Maps for easier referencing in the future.
    slaveIEDdevicecolumn = askForTextInput("Enter the column header for Slave IED Devices in the SCADA Map: ")
    # Replace any non-digit or non-letter symbols with underscores
    scadamapdf[slaveIEDdevicecolumn] = scadamapdf[slaveIEDdevicecolumn].str.strip().replace(r'[^a-zA-Z0-9]', '_', regex=True)
    # Add letters to the beginning of some devices to match Data Maps
    # Adding an L
    scadamapdf.loc[
        scadamapdf[slaveIEDdevicecolumn].str.match(r'^\d') &
        scadamapdf[slaveIEDdevicecolumn].str.upper().str.contains('LA'), # Checks if the device starts with a digit and contains 'LA'
        slaveIEDdevicecolumn
    ] = 'L' + scadamapdf[slaveIEDdevicecolumn]
    # Adding a K
    scadamapdf.loc[
        scadamapdf[slaveIEDdevicecolumn].str.match(r'^\d'), # Checks if the device starts with a digit
        slaveIEDdevicecolumn
    ] = 'K' + scadamapdf[slaveIEDdevicecolumn]

    # Change DataFrame column names to make future references easier
    # Binary Output Address
    matching_col = [
        col for col in scadamapdf.columns
        if all(p in str(col).lower() for p in ['binary', 'output', 'address'])
    ]
    if matching_col:
        old_name = matching_col[0]
        new_name = 'Binary Output Address'
        scadamapdf.rename(columns={old_name: new_name}, inplace=True)
    else:
        # Display popup window with error message
        tkinter.messagebox.showinfo("Error", "Binary Output Address column could not be found in SCADA Map. Check column header formatting.")
        sys.exit(1)
    # Binary DNP Address
    matching_col = [
        col for col in scadamapdf.columns
        if all(p in str(col).lower() for p in ['binary', 'dnp', 'address'])
    ]
    if matching_col:
        old_name = matching_col[0]
        new_name = 'Binary DNP Address'
        scadamapdf.rename(columns={old_name: new_name}, inplace=True)
    else:
        # Display popup window with error message
        tkinter.messagebox.showinfo("Error", "Binary DNP Address column could not be found in SCADA Map. Check column header formatting.")
        sys.exit(1)
    # Point Description column
    matching_col = [
        col for col in scadamapdf.columns
        if any(p in str(col).lower() for p in ['description', 'nomenclature'])
    ]
    if matching_col:
        old_name = matching_col[0]
        new_name = 'Point Description'
        scadamapdf.rename(columns={old_name: new_name}, inplace=True)
    else:
        # Display popup window with error message
        tkinter.messagebox.showinfo("Error", "Point Description column could not be found in SCADA Map. Check column header formatting.")
        sys.exit(1)
    # IED Wordbit column
    scadamapdf.rename(columns={wordbitcolumn: 'Relay Element'}, inplace=True)
    scadamapdf['Relay Element'] = scadamapdf['Relay Element'].str.strip()
    # Slave IED Device column
    scadamapdf.rename(columns={slaveIEDdevicecolumn: 'Slave IED Device'}, inplace=True)
    scadamapdf['Slave IED Device'] = scadamapdf['Slave IED Device'].str.strip()
    # IED DNP Index column
    iedDNPindexcolumn = askForTextInput("Enter the column header for the IED DNP Indices in the SCADA Map: ")
    scadamapdf.rename(columns={iedDNPindexcolumn: 'IED DNP Index'}, inplace=True)
    scadamapdf['IED DNP Index'] = scadamapdf['IED DNP Index'].str.strip()
    # Drop unnecessary columns
    scadamapdf = scadamapdf[['Binary Output Address', 'Binary DNP Address', 'Point Description', 'Slave IED Device', 'Relay Element', 'IED DNP Index']]

    # Some SCADA maps have single rows for both the .operTrip & .operClose operations. These should be split into their own individual rows for easier referencing in the future.
    new_rows = []

    for idx, row in scadamapdf.iterrows():
        col_splits = {}
        split_needed = False
        # Check both columns
        for col in ['Relay Element', 'IED DNP Index']:
            val = row[col]
            if isinstance(val, str):
                if col == 'IED DNP Index' and (',' in val or ':' in val):  # Split the second column
                    parts = val.split(',') if ',' in val else val.split(':')
                    col_splits[col] = parts
                    split_needed = True
                elif ':' in val or ':' in val:  # Split the first column
                    parts = val.split(':')
                    col_splits[col] = parts
                    split_needed = True
        if split_needed:
            original = row.copy()
            new = row.copy()
            for col in ['Relay Element', 'IED DNP Index']:
                if col in col_splits:
                    # Assign the first and second parts of each column to the two rows
                    original[col] = col_splits[col][0]
                    new[col] = col_splits[col][1]
            # Append the original row and the new row with same index
            new_rows.append((idx, original))
            new_rows.append((idx, new))
        else:
            # Keep row unchanged
            new_rows.append((idx, row))

    # Create a new DataFrame with the correct expanded rows, preserving the index
    scadamapdf = pd.DataFrame([row for _, row in new_rows], index=[idx for idx, _ in new_rows])
    return scadamapdf

def getDataMapFolder():
    """
    Opens a window to select a file directory

    Return:
    (str) - file directory path
    """
    Tk().withdraw()
    return tkinter.filedialog.askdirectory(title="Please select the folder containing all of your Data Maps.")

# Keeps track of whether the Data Maps use RTAC Tag Aliases
rtac_point_names = False

def convertDataMaps(dataMapFolder):
    """
    Finds all data maps in the selected directory and converts them into DataFrames.

    Parameter:
    dataMapFolder (str) - file directory path for Data Maps

    Return:
    (dict) - dictionary containing all converted Data Maps into DataFrames as values with the device names as the keys
    """

    # key = device name
    # value = DataFrame
    datamaps = {}

    for d in os.listdir(dataMapFolder):
        currentpath = dataMapFolder + "\\" + d
        # Check file name
        if ("data" in currentpath.lower()) and ("map" in currentpath.lower()) and (".xlsx" in currentpath.lower()):
            # Data maps with Control Points as Dataframe
            try:
                current = pd.read_excel(currentpath, sheet_name='Control Points')
            except:
                continue
            # Find device name
            found = False
            for row in range(current.shape[0]):
                for col in range(current.shape[1]):
                    cell_value = str(current.iat[row, col]).strip().upper()
                    if 'DEVICE NAME' in cell_value:
                        try:
                            device_name = str(current.iat[row + 1, col]).strip()
                            datamaps[device_name] = current
                            found = True
                            break
                        except Exception:
                            # Skip if device name is unreadable
                            break
                if found:
                    break
            # Parse data to extract only what is needed for SCADA map
            header_position = np.where(current == 'RELAY ELEMENT') # Find where data starts
            current.columns = current.iloc[header_position[0][0]]
            current = current[header_position[0][0] + 1:]
            current_parsed = pd.DataFrame(current['RELAY ELEMENT'])
            # Get a point name from RTAC
            if ('RTAC POINT NAME' in current.columns):
                current_parsed['POINT NAME'] = current['RTAC POINT NAME']
                rtac_point_names = True
            if ('RTU POINT NAME' in current.columns):
                current_parsed['POINT NAME'] = current['RTU POINT NAME']
                rtac_point_names = True
            # OR get a point name from HMI
            if ('HMI POINT NAME' in current.columns):
                current_parsed['POINT NAME'] = current['HMI POINT NAME']
                # Convert HMI point name to RTU-usable format
                current_parsed['POINT NAME'] = current_parsed['POINT NAME'].str.replace('.', '_', regex=False)
            current_parsed['POINT DESCRIPTION'] = current['POINT DESCRIPTION']
            # Get DNP Point Index
            matching_col = [
                col for col in current.columns
                if all(p in str(col).lower() for p in ['dnp', 'point', 'index'])
            ]
            current_parsed['DNP POINT INDEX'] = current[matching_col[0]]
            # Get whether the point is being sent to SCADA
            current_parsed['SCADA'] = current['SCADA']
            current_parsed['SCADA'] = current_parsed['SCADA'].notna() # Empty cells will evaluate to False, cells marked otherwise will evaluate to True
            datamaps[device_name] = current_parsed
    return datamaps

def writeStructuredText(scadamapdf, datamaps):
    """
    Writes the structured text to copy & paste into the RTAC

    Parameters:
    scadamapdf (DataFrame) -  SCADA Map as DataFrame
    datamaps (dict) - dictionary containing all Data Maps as DataFrames

    Return:
    (str) - structured text to copy & paste into the RTAC
    """

    structured_text = "\n"
    previous_device = "" # Keeps track of when a new device is used in the structured text
    operTrip = True # Keeps track of whether the point should include .operTrip or .operClose in the structured text

    UNIFORM_SPACING = [60, 42]

    # Iterate through every SCADA map entry
    for index, entry in scadamapdf.iterrows():
        # Get matching Data Map
        if entry['Slave IED Device'] not in datamaps:
            structured_text += f"\n DATA MAP FOR {entry['Slave IED Device']} WAS NOT FOUND.\n"
            continue
        else:
            current_datamap = datamaps.get(entry['Slave IED Device']) # The Slave IED Device value in the SCADA Map DataFrame should match the device name key for the dictionary of Data Maps.
        
        if (previous_device != "" and entry['Slave IED Device'] != previous_device):
            structured_text += "\n"
        previous_device = entry['Slave IED Device']
        if current_datamap is not None:
            # Get Data Map entry based on relay element
            current_point = current_datamap[current_datamap['RELAY ELEMENT'].str.fullmatch(re.escape(entry["Relay Element"]), na=False)] # Must escape special characters otherwise will return invalid.
            if not current_point.empty:
                current_point_name = current_point['POINT NAME'].iloc[0]

                # Check if marked for SCADA
                if not current_point['POINT NAME'].iloc[0]:
                    structured_text += f"\nWARNING: {current_point_name} IS NOT MARKED FOR SCADA IN DATA MAP.\n"

                current_point_index = current_point['DNP POINT INDEX'].iloc[0]

                # Check if indices match
                if int(current_point_index) != int(entry['IED DNP Index']):
                    structured_text += f"\n WARNING: THE INDICES LISTED IN THE SCADA MAP & DATA MAP DO NOT MATCH FOR {current_point_name}."

                # If the point name was derived from the HMI Point Name, the device name will need to be removed from the half portion of the Point Name returned from the DataFrame.
                if not rtac_point_names:
                    split_parts = current_point_name.split(entry['Slave IED Device'])
                    if len(split_parts) > 1:
                        current_point_name = split_parts[1].strip()
                current_point = current_point_name
                # Begin writing line(s) for structured text
                first_section = f"{entry['Slave IED Device']}_DNP.BO_{current_point_index:05}{current_point}"
                # Check operation
                if operTrip:
                    first_section += ".operTrip"
                else:
                    first_section += ".operClose"
                structured_text += f"{first_section:<{UNIFORM_SPACING[0]}}"
                second_section = f":= SCADA_DNP.BO_{index:05}"
                # Check operation
                if operTrip:
                    second_section += ".operTrip;"
                    operTrip = False
                else:
                    second_section += ".operClose;"
                    operTrip = True
                structured_text += f"{second_section:<{UNIFORM_SPACING[1]}}"
                structured_text += f"// {entry['Point Description']}\n"
    return structured_text

def main():
    scadamapdf = cleanSCADAMap(getSCADAMap())
    datamapfolder = getDataMapFolder()
    datamaps = convertDataMaps(datamapfolder)
    structured_text = writeStructuredText(scadamapdf, datamaps)
    # Output binary output programming to txt file
    output_file = "RTAC Binary Output Structured Text.txt"
    output_file = datamapfolder + "/RTAC Binary Output Structured Text.txt"
    with open(output_file, 'w') as ofile:
        ofile.write(structured_text)
    # Display popup window with completion message
    tkinter.messagebox.showinfo("", f"File saved at {output_file}")

if __name__ == "__main__":
    main()