#%% Convert a Folder of PowerPoint PPTs to PDFs

# Purpose: Converts all PowerPoint PPTs in a folder to Adobe PDF

# Original Author:  Matthew Renze
# Modified by: Sam Malyarov

# Note: Also works with PPTX file format

#%% Import libraries

import os
import comtypes.client

#%% Get console arguments
input_folder_path = os.getcwd()
output_folder_path = os.getcwd()

#%% Get files in input folder
input_file_paths = os.listdir(input_folder_path)

#%% Convert each file
for input_file_name in input_file_paths:

    # Skip if file does not contain a power point extension
    if not input_file_name.lower().endswith((".ppt", ".pptx")):
        continue
    
    # Create input file path
    input_file_path = os.path.join(input_folder_path, input_file_name)
        
    # Create powerpoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    
    # Set visibility to minimize
    powerpoint.Visible = 1
    
    # Open the powerpoint file
    slides = powerpoint.Presentations.Open(input_file_path)
    
    # Get base file name
    file_name = os.path.splitext(input_file_name)[0]
    
    # Create output file path
    output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
    
    # Save as PDF (formatType = 32)
    slides.SaveAs(output_file_path, 32)
    
    # Close the slide deck
    slides.Close()
    
print("Готово!")