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
    import pandas as pd
    import numpy as np
    import re
except:
    install("pandas")
    install("numpy")

def getSCADAMap():
    """
    Returns a DataFrame of the SCADA Map selected in the file dialog box

    Return: (DataFrame) DataFrame of the SCADA Map
    """
    Tk().withdraw()
    scadamappath = tkinter.filedialog.askopenfilename()
    return pd.read_excel(scadamappath, sheet_name='BINARY INPUTS')

def cleanSCADAMap(scadamapdf):
    """
    Performs data cleaning on a DataFrame for a SCADA Map. This includes removing unnecessary columns for SCADA programming, identifying Binary DNP Addresses, & altering the
    device name entries to match data maps.

    Parameter: scadamapdf (DataFrame) - DataFrame of the SCADA Map

    Return: scadamapdf (DataFrame) - cleaned DataFrame of the SCADA Map
    """
    # Drop entries with no Point Description
    scadamapdf = scadamapdf.dropna(subset=['Point Nomenclature/Description'])
    # Only include relevant columns
    scadamapdf = scadamapdf[['Binary\nDNP\nAddress', 'Point Nomenclature/Description', 'RTU\nInvert', 'Slave IED Device', 'Slave IED Wordbit', 'Slave IED DNP']]
    # Drop entries with no BI index
    scadamapdf = scadamapdf[scadamapdf['Binary\nDNP\nAddress'] != "-"]
    # Duplicate BI indexes to SCADA entries
    scadamapdf['Binary\nDNP\nAddress'] = scadamapdf['Binary\nDNP\nAddress'].ffill()
    # Drop entries with invalid BI index
    scadamapdf = scadamapdf[pd.to_numeric(scadamapdf['Binary\nDNP\nAddress'], errors='coerce').notnull()]
    # Clarify point inversion
    scadamapdf.loc[scadamapdf['RTU\nInvert'] == "X", 'RTU\nInvert'] = True
    scadamapdf['RTU\nInvert'] = scadamapdf['RTU\nInvert'].fillna(value=False)
    # Alter Slave IED Device to match data map device
    scadamapdf['Slave IED Device'] = scadamapdf['Slave IED Device'].apply(lambda x: re.sub(r'[^a-zA-Z0-9 ()]', '_', str(x)) if isinstance(x, str) or isinstance(x, float) else x)
    scadamapdf['Slave IED Device'] = scadamapdf['Slave IED Device'].apply(lambda x: "K" + x if (x[0].isdigit() and ("K" in x or "MO" in x)) else x)
    scadamapdf['Slave IED Device'] = scadamapdf['Slave IED Device'].apply(lambda x: "L" + x if x[0].isdigit() and "L" in x else x)
    return scadamapdf

def getDataMapFolder():
    """
    Returns the path of the file directory chosen in the file dialog box

    Return: (str) - path of the file directory chosen in the file dialog box
    """
    Tk().withdraw()
    return tkinter.filedialog.askdirectory()

def getDataMaps(datamapfolder):
    """
    Identifies the device name for each data map in the data map folder. Converts all data maps provided in a file directory to DataFrames. Once converted, DataFrames are added
    to a dictionary containing all data maps. The keys for each data map DataFrame are their corresponding device name provided in the data map. This function also performs data cleaning
    for each data map DataFrame.

    Parameter: datamapfolder (str) - path to folder of data maps

    Return: datamaps (dict) - dictionary consisting of all data map DataFrames with their corresponding device names as the keys
    """
    datamaps = {}
    # Make DataFrames for each data map
    for d in os.listdir(datamapfolder):
        currentpath = datamapfolder + "\\" + d
        if ("data" in currentpath.lower()) and ("map" in currentpath.lower()) and (".xlsx" in currentpath.lower()):
            current = pd.read_excel(currentpath)
            # Find device name
            device_name_row, device_name_col = np.where(current == 'DEVICE NAME')
            if device_name_row.size > 0 and device_name_col.size > 0:
                device_name_row += 1
                device_name = current.iat[device_name_row[0], device_name_col[0]].strip()  # Remove leading/trailing spaces
            else:
                continue
            # Parse data to extract only what is needed for SCADA map
            header_position = np.where(current == 'RELAY ELEMENT')
            current.columns = current.iloc[header_position[0][0]]
            current = current[header_position[0][0] + 1:]
            current_parsed = pd.DataFrame(current['RELAY ELEMENT'])
            if ('RTAC POINT NAME' in current.columns):
                current_parsed['POINT NAME'] = current['RTAC POINT NAME']
            if ('RTU POINT NAME' in current.columns):
                current_parsed['POINT NAME'] = current['RTU POINT NAME']
            current_parsed['POINT DESCRIPTION'] = current['POINT DESCRIPTION']
            datamaps[device_name] = current_parsed
    return datamaps

def writeStructuredText(scadamapdf, datamaps):
    """
    Writes the structured text for the SCADA programming in the RTAC.

    Parameters:
    scadamapdf (DataFrame) - cleaned DataFrame of the SCADA Map
    datamaps (dict) - dictionary consisting of all data map DataFrames with their corresponding device names as the keys

    Return: structured_text (str) - binary input SCADA programming, DOES NOT INCLUDE SEMICOLONS
    """
    # Begin writing SCADA programming structured text
    structured_text = "// ************************** SCADA DNP BINARY INPUTS **************************\n"
    structured_text +=  "// This program provides mapping of RTAC points to SCADA DNP Binary Inputs.\n\n"
    current_index = -1
    how_to_add = "OR"
    # Iterate through every SCADA map entry
    for index, entry in scadamapdf.iterrows():
        current_binary = int(entry["Binary\nDNP\nAddress"])
        if current_binary != current_index:
            # Add group comment, time, validity etc.
            structured_text += f"\n// {entry['Point Nomenclature/Description']}\n"
            structured_text += f"SCADA_DNP.BI_{current_binary:05}.t := SYS_TIME();\n"
            structured_text += f"SCADA_DNP.BI_{current_binary:05}.q.validity := GOOD;\n"
            structured_text += f"SCADA_DNP.BI_{current_binary:05}.stVal	:=\n"
            current_index = current_binary
            first_expression = True
        # Add logic expression(s)
        if pd.notna(entry['Slave IED Device']):
            # Note how the expression(s) are conjucated
            if "OR" in entry['Slave IED Device']:
                how_to_add = "OR"
                first_expression = True
            elif "AND" in entry['Slave IED Device']:
                how_to_add = "AND"
                first_expression = True
            elif "Grouped as" in entry['Slave IED Device']:
                how_to_add = ""
                first_expression = True
            else:
                # Get matching data map
                current_datamap = datamaps.get(entry['Slave IED Device'])
                if current_datamap is not None:
                    # Get data map entry based on relay element
                    current_point = current_datamap[current_datamap['RELAY ELEMENT'].str.contains(entry["Slave IED Wordbit"])]
                    if not current_point.empty:
                        current_point = current_point['POINT NAME'].iloc[0]
                        # Add line to logic
                        structured_text += "			"
                        # Write remainder of line
                        if first_expression:
                            # Check if point should be inverted
                            if entry['RTU\nInvert']:
                                structured_text += "NOT "
                            structured_text += f"{current_point}.stVal								//	{entry['Point Nomenclature/Description']}   {entry['Slave IED Device']}   {entry['Slave IED Wordbit']}\n"
                            first_expression = False
                        else:
                            structured_text += f"{how_to_add} "
                            # Check if point should be inverted
                            if entry['RTU\nInvert']:
                                structured_text += "NOT "
                            structured_text += f"{current_point}.stVal								//	{entry['Point Nomenclature/Description']}   {entry['Slave IED Device']}   {entry['Slave IED Wordbit']}\n"
                    else:
                        structured_text += f"// WARNING: NO MATCHING RELAY ELEMENT FOUND FOR {entry['Slave IED Device']}\n"
                else:
                    structured_text += f"// WARNING: NO MATCHING DATA MAP FOUND FOR {entry['Slave IED Device']}\n"
    return structured_text

def main():
    scadamapdf = cleanSCADAMap(getSCADAMap())
    datamapfolder = getDataMapFolder()
    datamaps = getDataMaps(datamapfolder)
    structured_text = writeStructuredText(scadamapdf, datamaps)
    # Output binary input programming to txt file
    output_file = "RTAC Binary Input Structured Text.txt"
    output_file = datamapfolder + "/RTAC Binary Input Structured Text.txt"
    with open(output_file, 'w') as ofile:
        ofile.write(structured_text)
    # Display popup window with completion message
    tkinter.messagebox.showinfo("", f"File saved at {output_file}")

if __name__ == "__main__":
    main()